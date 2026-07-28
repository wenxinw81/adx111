# 打包为 Mac App / Windows EXE

## 版本

Mac 会生成两个版本：

- `Direct`：默认直连 Doris，给公司内能直连数据库的同事使用。
- `SSH`：默认走 SSH 隧道，给本机这种需要跳板访问 Doris 的环境使用。

Windows 只生成 `Direct` 直连版。

## Mac

在 Mac 上运行：

```bash
./build_mac_app.sh
```

产物：

```text
dist/ADXReportAgent-Direct-macOS.zip
dist/ADXReportAgent-SSH-macOS.zip
```

本机使用：

```text
ADXReportAgent-SSH.app
```

给能直连数据库的同事使用：

```text
ADXReportAgent-Direct.app
```

首次启动后会在用户目录创建：

```text
~/ADXReportAgent/.env
~/ADXReportAgent/outputs
```

`.env` 默认已经写入公司 Doris 连接信息；SSH 版本也会读取备用 SSH 信息。需要改端口、密码或映射表时再改这个文件。

Mac App 是本地临时构建，未做企业开发者签名。首次打开如果提示安全限制，右键点击 App，选择“打开”，确认一次即可。

## Windows

Windows EXE 建议用 GitHub Actions 生成。把项目推到 GitHub 后，在 Actions 里运行 `Build Windows EXE`，产物在 workflow artifact：

```text
ADXReportAgent-Direct-windows.zip
```

给能直连数据库的 Windows 同事：

```text
ADXReportAgent-Direct.exe
```

Windows 首次启动后会在用户目录创建：

```text
C:\Users\用户名\ADXReportAgent\.env
```

Windows 只支持直连版；如果某台 Windows 不能直连数据库，先让网络/数据库访问打通后再使用。
