# ADX 日报 Agent

一句话生成广告投放基础分析报表。

## Agent 架构

当前项目已按 `agent-dev-standard` 规范改成 LangGraph 编排：

```text
parse_request → plan_report → execute_report → validate_report → reflect / end
```

核心目录：

- `adx_report_agent/state.py`：图状态、校验结果、reflection 决策
- `adx_report_agent/graph.py`：LangGraph `StateGraph`
- `adx_report_agent/nodes/`：parse、plan、execute、validate、reflect 节点
- `adx_report_agent/tools/`：工具注册、报表执行工具、Excel 校验工具
- `adx_report_agent/llm_client.py`：统一 LLMClient 预留接口
- `adx_report_agent/prompts.py`：后续 LLM 结论生成 prompt 模板

节点只返回 state patch；脚本执行和 Excel 校验通过 tool registry 调用；失败统一进入 reflection 节点。

## 运行

## 离线安装

如果对方电脑不能访问 PyPI，但已经安装 Python 3.12，可以使用离线包里的脚本：

Mac：

```bash
cp env.example .env
./install_offline_mac.sh
./run_adx_web.sh
```

Windows：

```bat
copy env.example .env
install_offline_windows.bat
run_adx_web.bat
```

离线依赖只覆盖：

- macOS Apple Silicon / ARM64 + Python 3.12
- Windows 64 位 + Python 3.12

## 数据库信息填在哪里

推荐复制 `env.example` 为 `.env`，在 `.env` 里填写数据库信息：

```bash
cp env.example .env
```

Mac/Linux 的 `.env` 示例：

```text
DORIS_HOST=192.168.100.23
DORIS_PORT=29030
DORIS_DATABASE=ads
DORIS_USER=WishFox
DORIS_PASSWORD=填写数据库密码
ADX_SSH_ENABLED=false
```

Windows 也同样复制 `env.example` 为 `.env`，用记事本改里面的值即可。

也可以直接改 `configs/agent.direct.example.json` 里的 `database.host`、`database.port`、`database.database`、`database.user`。密码不要写进 JSON，放在 `.env` 的 `DORIS_PASSWORD`。

### Web 运行

Mac/Linux 直连数据库：

```bash
python3 -m pip install -e .
cp env.example .env
./run_adx_web.sh --config configs/agent.direct.example.json --host 0.0.0.0 --port 8787
```

Windows 直连数据库：

```bat
py -m pip install -e .
copy env.example .env
run_adx_web.bat
```

打开：

```text
http://127.0.0.1:8787
```

如果让局域网其他人访问，把 `127.0.0.1` 换成运行 Web 服务那台电脑的 IP。

Web 页面支持输入：

```text
看下昨天数据
看下昨天花销
看下昨天竞价
帮我生成昨天的218订单的基础数据
帮我生成昨天218订单花销
帮我生成昨天order_id=218竞价
看下 2026-07-25 数据
```

当前自然语言解析是确定性规则解析，支持相对日期、显式日期和三类报表关键词；如果后续要支持更自由的表达，可以把 `adx_report_agent/llm_client.py` 接到云端模型，让模型输出结构化 JSON 后再进入 LangGraph。

APP 包名映射表默认读取 `data/app_mapping.csv`。如果后续要替换成新映射表，可以覆盖这个文件，或通过环境变量指定：

```bash
export APP_MAPPING_CSV='/path/to/new-dsp-app包名.csv'
```

### 对话式运行

当前电脑需要 SSH 隧道时：

```bash
export ADX_SSH_PASSWORD='...'
export DORIS_PASSWORD='...'
./run_adx_agent.sh
```

别人电脑可以直连数据库时：

```bash
export DORIS_PASSWORD='...'
./run_adx_agent.sh --config configs/agent.direct.example.json
```

启动后可以输入：

```text
看下昨天数据
看下 2026-07-25 数据
看下 2026-07-25 花销
看下昨天花费
看下 2026-07-25 竞价
看下昨天出价分析
帮我生成昨天的218订单的基础数据
帮我生成昨天218订单花销
帮我生成昨天order_id=218竞价
改标准：APP Top50
改标准：小时分析不要展示无出价小时
显示标准
退出
```

### 单条命令运行

当前电脑需要 SSH 隧道时：

```bash
export ADX_SSH_PASSWORD='...'
export DORIS_PASSWORD='...'
export APP_MAPPING_CSV='/path/to/dsp-app包名.csv'
./run_adx_report.sh "看下昨天数据"
```

别人电脑可以直连数据库时：

```bash
export DORIS_PASSWORD='...'
export APP_MAPPING_CSV='/path/to/dsp-app包名.csv'
./run_adx_report.sh "看下昨天数据" --config configs/agent.direct.example.json
```

指定日期：

```bash
python3 -m adx_report_agent.cli "看下 2026-07-25 数据"
```

生成花销专门分析：

```bash
./run_adx_report.sh "看下 2026-07-25 花销"
```

生成竞价专门分析：

```bash
./run_adx_report.sh "看下 2026-07-25 竞价"
```

## 给别人使用

把这些文件夹/文件一起发给对方：

- `adx_report_agent/`
- `scripts/`
- `configs/`
- `README_adx_report_agent.md`
- `.vendor_pymysql/`（如果对方不方便安装依赖）

对方电脑如果能联网，也可以不带 `.vendor_pymysql/`，让对方执行：

```bash
python3 -m pip install pandas openpyxl pydantic pymysql
```

数据库密码不要写进代码或配置文件，通过环境变量传：

```bash
export DORIS_PASSWORD='数据库密码'
```

如果对方电脑能直连数据库，使用 `configs/agent.direct.example.json`；如果也需要跳板 SSH，使用 `configs/agent.runtime.example.json`，并设置 `ADX_SSH_PASSWORD`。

## 修改标准

后续优先改 `configs/basic_report.json`：

- `sheet_order`：调整 sheet 顺序
- `rules.drop_zero_bid_hours`：小时分析是否删除无出价小时
- `rules.landing_page_requires_click`：落地页曝光/跳转是否必须先经过 click
- `rules.track_join_rule`：埋点关联口径，目前跳转优先从 URL 的 `rid/pyck` 关联点击
- `rules.app_bundle_to_app_name`：是否用包名映射表展示 APP 名称
- `style`：统一颜色

花销专门分析标准在 `configs/spend_report.json`。如果要新增“竞价专门分析”，再增加新的标准文件和 runner 分支。
竞价专门分析标准在 `configs/bidding_report.json`，当前包含 APP、时段、订单策略素材、出价分层，以及四个维度两两交叉的 6 张表。
