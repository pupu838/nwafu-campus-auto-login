# NWAFU Campus Network Auto Login (Windows + Edge)

## 中文说明

### 1. 项目简介
这是一个面向 Windows 的 NWAFU 校园网自动认证工具（Tkinter GUI）。
程序会持续检测当前 Wi-Fi，当检测到连接到 `NWAFU` 时，自动打开认证页并完成登录流程。

### 2. 主要功能
- 自动监控当前 SSID，仅在 `NWAFU` 下触发自动认证。
- Selenium + Edge 自动登录：
  - 若页面已出现“注销”元素，判定为已登录，跳过提交。
  - 若未登录，自动识别账号/密码输入框并点击“登录”。
- 认证保活检测：
  - 已登录状态下按间隔探测连通性；
  - 若判定离线，立即触发重连。
- 凭据安全存储：
  - 账号密码保存到 Windows Credential Manager（`keyring`），不写入明文配置文件。
- 托盘与窗口行为：
  - 可收至系统托盘继续运行；
  - 支持“点击 X 直接退出”或“点击 X 收至托盘”两种行为。
- 开机自启与自修复：
  - 优先写入注册表 Run；
  - 失败时回退到启动文件夹脚本；
  - 启动时自动检查并修复失效/重复启动项。
- 单实例运行保护：
  - 使用系统互斥体 + 锁文件，避免重复启动。

### 3. 运行环境
- 操作系统：Windows 10 / 11
- 浏览器：Microsoft Edge（已安装）
- Python：建议 3.9+（源码运行）

### 4. 安装与运行

#### 4.1 安装依赖
```powershell
python -m pip install -r requirements.txt
```

#### 4.2 源码运行（推荐无控制台）
```powershell
pythonw nwafu_login.py
```

如需看到控制台输出，可用：
```powershell
python nwafu_login.py
```

#### 4.3 启动参数
- `--autostart`：标记为开机自启场景启动。
- `--minimized`：启动后收至托盘/最小化。

示例：
```powershell
pythonw nwafu_login.py --autostart --minimized
```

### 5. 首次使用建议流程
1. 启动程序。
2. 在“凭据”区域输入账号与密码，点击“保存凭据”。
3. 在“设置”区域按需调整：
   - 轮询间隔（默认 5 秒）
   - 认证保活检测间隔（默认 60 秒）
   - 开机自启 / 关闭窗口行为 / 认证成功后自动收托盘
4. 点击“开始监控”。
5. 可用“立即检测”手动触发一次认证流程。

### 6. 核心行为说明

#### 6.1 自动认证触发逻辑
- 仅自动模式下要求当前 SSID 为 `NWAFU`。
- 手动“立即检测”会强制执行，不受 SSID 限制。
- 登录失败后会在 60 秒后自动重试。

#### 6.2 认证保活检测
- 仅在当前会话已认证时执行。
- 默认每 60 秒探测一次 `msftconnecttest` 页面。
- 如果被重定向到认证门户，判定为离线并立即重连。

#### 6.3 Selenium 导航回退
- 主策略：浏览器内 JS 跳转到门户页面。
- 若仍停留 `about:blank`，再尝试跳转探测页。
- 仍失败时触发“本地 Edge 回退”打开页面（自动场景有冷却时间，避免频繁拉起）。

### 7. 配置、日志与数据路径

所有运行时文件默认在 `%LOCALAPPDATA%\NWAFUAutoLogin\` 下。

- 配置文件：`%LOCALAPPDATA%\NWAFUAutoLogin\config.json`
- 启动诊断日志：`%LOCALAPPDATA%\NWAFUAutoLogin\startup.log`
  - 超过 1MB 自动截断
- Edge 独立用户目录：`%LOCALAPPDATA%\NWAFUAutoLogin\edge-profile`
- 单实例锁文件：`%LOCALAPPDATA%\NWAFUAutoLogin\instance.lock`

凭据不在 `config.json`，而在 Windows Credential Manager 内，服务名为 `NWAFUAutoLogin`。

#### 7.1 `config.json` 字段（默认值）
- `target_ssid`: `"NWAFU"`（代码强制固定）
- `portal_url`: `https://portal.nwafu.edu.cn/srun_portal_success?ac_id=1&theme=pro`（代码强制固定）
- `poll_interval_sec`: `5`
- `auth_probe_interval_sec`: `60`
- `login_timeout_sec`: `25`
- `edge_profile_dir`: `%LOCALAPPDATA%\NWAFUAutoLogin\edge-profile`
- `keep_browser_on_failure`: `true`
- `autostart_enabled`: `false`
- `exit_on_close`: `true`
- `auto_hide_to_tray_on_success`: `true`

### 8. 开机自启机制
- 注册表路径：`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`
- 启动文件夹脚本：`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\NWAFUAutoLogin.vbs`

策略说明：
- 启用自启时优先使用注册表 Run。
- 若注册表写入失败，回退到启动文件夹脚本。
- 若检测到两个通道同时存在，会自动清理为单通道，避免重复启动。

自动监控延迟：
- 带 `--autostart` 启动：30 秒后开始监控（降低开机阶段争用）。
- 普通启动但配置要求自动监控：约 500ms 后开始监控。

### 9. 打包 EXE

#### 9.1 Onefile
```powershell
pyinstaller --onefile --windowed --name NWAFUAutoLogin nwafu_login.py
```

#### 9.2 Onedir
```powershell
pyinstaller --onedir --windowed --name NWAFUAutoLogin nwafu_login.py
```

已包含示例 spec 文件：`NWAFUAutoLogin.spec`。

### 10. 常见问题排查

#### 10.1 开机后没有自动运行
1. 查看 `%LOCALAPPDATA%\NWAFUAutoLogin\startup.log`。
2. 检查 Run 项是否存在 `NWAFUAutoLogin`。
3. 检查启动文件夹脚本是否存在。
4. 手动启动一次程序，观察是否出现“启动项修复”日志。

#### 10.2 一直提示登录失败
1. 点击“打开认证页”确认校园网门户可正常访问。
2. 确认账号密码正确并重新保存凭据。
3. 查看日志是否出现“未找到登录按钮/输入框”。
4. 若页面结构改版，可临时保留失败浏览器窗口进行人工确认（`keep_browser_on_failure=true`）。

#### 10.3 托盘不可用
- 可能缺少 `pystray` 或 `Pillow`，程序会自动回退为最小化到任务栏。

### 11. 项目文件
- `nwafu_login.py`：主程序（GUI、监控、认证、托盘、自启、凭据管理）
- `requirements.txt`：运行与打包依赖
- `NWAFUAutoLogin.spec`：PyInstaller 打包配置示例

### 12. 安全提示
- 本工具会自动提交校园网账号密码，请仅在可信设备上使用。
- 建议为系统账号设置登录密码并启用磁盘加密（如 BitLocker）。

---

## English Guide

### 1. Overview
This is a Windows desktop auto-login utility for NWAFU campus network.
It continuously monitors Wi-Fi status and triggers the login workflow when connected to `NWAFU`.

### 2. Key Features
- Monitors current SSID and only auto-triggers on `NWAFU`.
- Selenium + Edge login automation:
  - Detects already-authenticated state via logout elements.
  - Fills username/password and clicks login when needed.
- Session health probe:
  - Periodically checks connectivity after successful login.
  - Immediately re-authenticates when session is considered offline.
- Secure credential storage via Windows Credential Manager (`keyring`).
- Tray behavior support:
  - Minimize to tray and keep monitoring in background.
  - Configurable close behavior (exit vs. hide).
- Autostart with self-healing:
  - Prefers Registry Run entry.
  - Falls back to Startup folder script if needed.
  - Repairs invalid/duplicate startup entries at app launch.
- Single-instance protection to prevent duplicate runs.

### 3. Requirements
- Windows 10/11
- Microsoft Edge installed
- Python 3.9+ (for source run)

### 4. Install and Run

#### 4.1 Install dependencies
```powershell
python -m pip install -r requirements.txt
```

#### 4.2 Run from source (no console window)
```powershell
pythonw nwafu_login.py
```

Optional (with console):
```powershell
python nwafu_login.py
```

#### 4.3 CLI arguments
- `--autostart`: marks a boot/autostart launch.
- `--minimized`: starts minimized to tray/taskbar.

Example:
```powershell
pythonw nwafu_login.py --autostart --minimized
```

### 5. Recommended First-Time Setup
1. Launch the app.
2. Save your campus account/password in the Credentials section.
3. Adjust settings as needed:
   - Poll interval (default: 5s)
   - Auth probe interval (default: 60s)
   - Autostart / close behavior / auto-hide-to-tray
4. Click `Start Monitoring`.
5. Use `Check Now` for a manual immediate check.

### 6. Behavior Details

#### 6.1 Auto-login trigger
- In auto mode, login runs only when SSID is `NWAFU`.
- Manual check bypasses SSID guard and runs immediately.
- Failed login is retried after 60 seconds.

#### 6.2 Session probe
- Runs only when current session is marked authenticated.
- Checks `msftconnecttest` endpoints at configured interval.
- If redirected to the portal host, session is considered offline and relogin is triggered.

#### 6.3 Navigation fallback
- Primary navigation: JS redirect to portal URL in automated Edge.
- If still `about:blank`, tries a probe URL.
- If navigation still fails, opens local Edge fallback flow (with cooldown in auto mode).

### 7. Paths and Data
Default runtime directory: `%LOCALAPPDATA%\NWAFUAutoLogin\`

- Config: `%LOCALAPPDATA%\NWAFUAutoLogin\config.json`
- Startup diagnostics log: `%LOCALAPPDATA%\NWAFUAutoLogin\startup.log` (rotates at 1MB)
- Edge profile: `%LOCALAPPDATA%\NWAFUAutoLogin\edge-profile`
- Instance lock: `%LOCALAPPDATA%\NWAFUAutoLogin\instance.lock`

Credentials are stored in Windows Credential Manager under service name `NWAFUAutoLogin`, not in `config.json`.

### 8. Autostart Implementation
- Registry key: `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`
- Startup folder script: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\NWAFUAutoLogin.vbs`

Policy:
- Registry Run is preferred.
- Startup-folder script is fallback.
- Duplicate channels are normalized to a single valid channel.

Monitoring startup delay:
- With `--autostart`: monitoring starts after 30s.
- Normal launch with auto-monitor enabled: starts after ~500ms.

### 9. Build EXE with PyInstaller

#### 9.1 Onefile
```powershell
pyinstaller --onefile --windowed --name NWAFUAutoLogin nwafu_login.py
```

#### 9.2 Onedir
```powershell
pyinstaller --onedir --windowed --name NWAFUAutoLogin nwafu_login.py
```

Reference spec file: `NWAFUAutoLogin.spec`.

### 10. Troubleshooting

#### 10.1 Not launching on boot
1. Check `%LOCALAPPDATA%\NWAFUAutoLogin\startup.log`.
2. Verify Registry Run entry exists.
3. Verify Startup folder script exists.
4. Manually run the app once and check whether startup-entry repair is logged.

#### 10.2 Login keeps failing
1. Open portal manually to verify portal reachability.
2. Re-save credentials and verify account/password.
3. Check logs for missing login button/inputs.
4. Keep failure browser open for debugging (`keep_browser_on_failure=true`).

#### 10.3 Tray icon unavailable
- Usually caused by missing `pystray`/`Pillow`; app will gracefully fallback to taskbar minimize.

### 11. Files
- `nwafu_login.py`: main app (GUI, monitor, login, tray, autostart, credentials)
- `requirements.txt`: dependencies
- `NWAFUAutoLogin.spec`: PyInstaller spec example

### 12. Security Notes
- The app automates campus credential submission. Use only on trusted devices.
- It is recommended to protect your Windows account and enable disk encryption.
