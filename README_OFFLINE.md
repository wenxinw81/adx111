# ADX 日报 Agent 离线安装说明

本包适用于已安装 Python 3.12 的电脑。

## Mac

```bash
cd adx_report_agent_web_offline_py312_win_mac
cp env.example .env
./install_offline_mac.sh
./run_adx_web.sh
```

## Windows

```bat
cd adx_report_agent_web_offline_py312_win_mac
copy env.example .env
install_offline_windows.bat
run_adx_web.bat
```

离线安装脚本会在当前目录创建 `.venv`，依赖安装在 `.venv` 里。

## 数据库信息

编辑 `.env`：

```text
DORIS_HOST=192.168.100.23
DORIS_PORT=29030
DORIS_DATABASE=ads
DORIS_USER=WishFox
DORIS_PASSWORD=数据库密码
ADX_SSH_ENABLED=false
```

启动后打开：

```text
http://127.0.0.1:8787
```

如果局域网其他人访问，用运行服务那台电脑的 IP：

```text
http://运行服务的电脑IP:8787
```
