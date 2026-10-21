# 邮件助手｜Mail Assistant

一款本地运行的桌面邮件提醒与分类工具，支持 Windows 和 macOS。它把多个支持 IMAP SSL 的邮箱汇总到轻量悬浮卡片中：每日固定时间检查当日新邮件、识别优先级、显示正文摘要，并可在应用内直接阅读原邮件正文。

> **目标：** 每天早上先看值得处理的事项，而不是重新翻遍所有邮箱。
>
> **隐私：** 邮件在本机读取与分类，不调用云端 AI。客户端密码/授权码存入 Windows Credential Manager 或 macOS Keychain；最近晨报只在本机保存必要元数据和最多 220 字的正文摘要。

## 功能

- **一眼识别来源：** 每封邮件卡片标出邮箱账号/来源、发件人、主题、时间、分类和优先级。
- **直接查看正文：** 点击悬浮卡片中的邮件会打开应用内正文阅读器，而不是只打开邮箱首页；正文按需从邮箱只读读取。阅读器另有“在网页版打开”入口。
- **仅提醒当日新邮件：** 桌面卡片和晨报列表只显示本地日历日期内的新邮件，不把过去日期的邮件混入今日提醒。
- **更适合晨间分拣：** 分类为“需尽快处理、日程与会议、课程与学术、求职与实习、账单与账户、通知与订阅、其他往来”，并额外标记“高 / 中 / 普通”优先级。
- **扩展邮箱：** 提供 Gmail、QQ、清华大学、Outlook/Hotmail、Yahoo、iCloud、网易 163/126、新浪、腾讯企业邮箱、阿里云企业邮箱、Zoho、Fastmail 等配置预设；也可以添加任意 IMAP SSL 邮箱。
- **更清爽、易读：** 浅色卡片、明确的层级、系统字体、高 DPI 感知；邮件较多时可滚动浏览。
- **定时通知：** 默认每天 08:00 检查。Windows 登录时自动运行；macOS 安装后会注册登录启动任务，并使用系统通知提醒。
- **只读工作方式：** 不发送、删除、移动或主动标记邮件为已读；不下载附件到本地。点击阅读时会从服务器提取可显示的正文文字。

## 支持的邮箱与登录方式

| 邮箱预设 | IMAP SSL 主机 | 说明 |
|---|---|---|
| Gmail | `imap.gmail.com:993` | 使用应用专用密码；当前版本尚不支持 Gmail OAuth 登录 |
| QQ 邮箱 | `imap.qq.com:993` | 先启用 IMAP，再使用授权码 |
| 清华大学邮箱 | `mails.tsinghua.edu.cn:993` | 按学校设置使用客户端专用密码 |
| Outlook / Hotmail | `outlook.office365.com:993` | 仅适用于账号允许 IMAP 密码/应用密码认证的情形 |
| Yahoo Mail | `imap.mail.yahoo.com:993` | 通常需要应用专用密码 |
| iCloud Mail | `imap.mail.me.com:993` | 使用 Apple 账号的 App 专用密码 |
| 网易 163 / 126 | `imap.163.com:993` / `imap.126.com:993` | 启用 IMAP 并生成客户端授权码 |
| 新浪邮箱 | `imap.sina.com:993` | 以账号安全设置为准 |
| 腾讯企业邮箱 | `imap.exmail.qq.com:993` | 需管理员允许 IMAP |
| 阿里云企业邮箱 | `imap.qiye.aliyun.com:993` | 企业服务器可能因组织配置而不同 |
| Zoho Mail | `imap.zoho.com:993` | 主机可能因数据中心区域不同而变化 |
| Fastmail | `imap.fastmail.com:993` | 使用应用专用密码 |
| 网易 yeah.net | `imap.yeah.net:993` | 启用 IMAP 并生成客户端授权码 |
| Foxmail 邮箱 | `imap.qq.com:993` | Foxmail.com 通常使用 QQ 邮箱服务器与授权码 |
| 中国移动 139 邮箱 | `imap.139.com:993` | 按 139 邮箱设置启用 IMAP 与客户端授权码 |

在设置窗口点 **“添加邮箱…”** 可新增其他预设账号，也可以选 **“自定义 IMAP 邮箱”**，填写显示名称、IMAP SSL 主机、端口、邮箱地址、客户端密码/授权码和可选的网页版地址。

**认证方式限制：** 本版本使用 IMAP SSL + 客户端密码/授权码，不提供 OAuth2 登录流程。部分 Gmail、Outlook/Microsoft 365、学校或企业账号会强制 OAuth2 或阻止基本密码认证；这些账号即使能填写主机，也可能无法连接。不要将普通网页登录密码用于需要应用专用密码的邮箱。Proton Mail 等需要专用 Bridge、特殊 TLS 或 OAuth2 的服务不保证可用。

## Windows 安装（Windows 10/11）

### 第一次安装或升级

1. 安装 Python 3.11 或更新版本，并启用 **Add Python to PATH**。
2. 打开本仓库：<https://github.com/zishengliu/mail-briefing-assistant>，点击 **Code → Download ZIP**，解压。
3. 进入解压后的项目文件夹，在文件资源管理器顶部地址栏输入 `powershell` 并回车。
4. 在 PowerShell 中逐条运行以下命令：

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\install_windows.ps1
   ```

   如果询问是否更改执行策略，输入 `Y`。安装脚本会复制程序、安装依赖、注册登录启动任务，并在开始菜单和桌面创建“邮件助手”快捷方式。
5. 以后从桌面的 **邮件助手** 图标或开始菜单打开即可。日常使用不需要再打开 PowerShell。

> 升级时请下载最新 ZIP 并重新运行安装脚本。程序会安装到 `%LOCALAPPDATA%\MailBriefingAssistant`，邮箱设置和凭据保存在独立的本机位置，不会因为升级而自动删除。

## macOS 安装（macOS 12 或更新版本建议）

1. 安装 Python 3.11 或更新版本。推荐使用 python.org 的 macOS 安装包，并确认它包含 Tcl/Tk 支持（运行 `python3 -c 'import tkinter'` 可检查）。
2. 从仓库下载 ZIP 并解压。
3. 打开 Terminal，切换到解压后的项目文件夹。例如解压在 Downloads 时：

   ```bash
   cd ~/Downloads/mail-briefing-assistant-main
   bash install_macos.sh
   ```

4. 安装脚本会创建虚拟环境、安装依赖，并在 `~/Applications` 创建名为 **邮件助手** 的应用；同时注册登录启动项。
5. 之后从 Applications 打开“邮件助手”，或把它拖到 Dock。日常使用不需要再次打开 Terminal。

macOS 的数据目录位于 `~/Library/Application Support/MailAssistant`，凭据由 macOS Keychain 管理。若 macOS 阻止应用启动，请确认这是你从自己的源码仓库下载并运行安装脚本生成的应用，再按系统安全提示允许启动。

## 配置邮箱

1. 打开 **邮件助手 - 设置**（Windows 开始菜单中可找到；macOS 可从应用菜单或本程序的设置按钮打开）。
2. 默认显示 Gmail、QQ 邮箱和清华大学邮箱。填写完整邮箱地址与客户端密码/授权码并勾选“启用”。
3. 需要添加其他服务时，点击 **添加邮箱…**，从预设列表选择服务，核对 IMAP 主机/端口、填写邮箱地址与授权信息。可选填写“IMAP 登录用户名”；默认使用完整邮箱地址，但 iCloud 等服务可能要求仅填写 @ 前面的用户名。自定义邮箱请填写服务商提供的 IMAP SSL 信息。
4. 点击 **保存设置**，然后点击 **测试邮箱连接**。确认连接成功后，点击“立即检查新邮件”即可手动测试。
5. 输入密码/授权码时不要使用普通网页登录密码，除非你的邮箱服务商明确允许。请勿把任何真实凭据发到聊天或公开 Issue。

常见客户端授权入口：

- Gmail：[Google 账号安全设置](https://myaccount.google.com/security) · [应用专用密码](https://myaccount.google.com/apppasswords)
- QQ 邮箱：[QQ 邮箱网页版](https://mail.qq.com/) → 设置 → 账户 → 开启 IMAP 并生成授权码
- 清华大学邮箱：[清华大学邮箱](https://mails.tsinghua.edu.cn/) → 按学校邮箱设置启用客户端访问
- iCloud：[Apple 账号](https://account.apple.com/) → 登录与安全 → App 专用密码（按账号当前界面为准）

## 日常使用

### 每日提醒时间
在设置窗口的“每日提醒”中填写 24 小时制时间，例如 `08:00` 或 `07:30`，然后保存。默认每天 08:00。程序会在每天设定时间检查邮件，更新桌面卡片，并通过操作系统通知提示当日新邮件与分类摘要。

**限制：** 电脑需要开机、用户已登录，而且邮件助手进程处于运行状态；电脑关机或用户完全退出登录时不能联网检查。Windows 通过登录启动任务拉起插件；macOS 使用 LaunchAgent。若错过当天时间，登录并启动程序后会尝试补跑当日检查。

### 阅读邮件
- 悬浮卡片收起时显示当日新邮件数量；点击“详情”展开。
- 每封邮件卡片显示邮箱来源、发件人、主题、时间、分类、优先级和简短正文摘要。
- **点击邮件卡片会打开应用内正文阅读器。** 正文在点击时通过只读 IMAP 从原邮箱读取；阅读器内可选择“在网页版打开”作为后备。
- 双击设置窗口中的邮件列表项也会打开正文阅读器。
- 正文阅读器不显示附件，不会标记邮件为已读。为避免无限下载，读取时最多取邮件正文前 2 MB，并最多显示 40,000 个可提取字符；纯图片、加密邮件或不支持的正文格式可能无法展示为文本。

### 当日邮件与分类
- 只显示邮件日期换算到本地时区后属于今天的邮件。昨天或更早的邮件不会混入今日晨报。
- 分类和优先级使用本地关键词规则，不是 AI 语义判断。涉及重要期限、资金、安全或法律事务时，请以原邮件内容为准。
- “高优先级”通常由紧急用语、截止时间和需回复线索触发；“中”用于部分日程、邀请、面试或付款线索。误分类时可以打开原邮件确认。

## 隐私与本地数据

- 邮箱密码/授权码保存在 Windows 系统凭据库或 macOS Keychain，不写入 Git 仓库。
- 本地设置、同步状态、日志和晨报元数据存放在用户目录；Windows 为 `%APPDATA%\MailBriefingAssistant`，macOS 为 `~/Library/Application Support/MailAssistant`。
- 为显示最近晨报，程序会在本地保存发件人、主题、邮箱来源、日期、分类、优先级和最多 220 字的正文摘要。摘要可能包含个人信息，请按敏感数据妥善保护。
- 完整正文仅在打开正文阅读器时临时读取，不保存到晨报缓存；附件不下载。邮件内容不会上传到 GitHub 或云端 AI。
- 应用使用只读 IMAP 访问收件箱，不发送、删除或移动邮件。

## 故障排查

**正文摘要为空**
- 新版会优先提取 `text/plain`，若邮件只有 HTML 则尝试抽取可见文字，并忽略脚本、样式与附件。
- 更新前的缓存会尝试重新读取原邮件来补齐摘要；若邮件已被删除、UID 已变化、账号连接失败、邮件加密或正文只包含图片，仍可能没有摘要。点击邮件卡片可尝试从服务器直接读取正文。
- 超大邮件只读取正文前 2 MB；某些嵌入式附件或特殊 MIME 结构可能无法完整提取。

**邮箱无法连接**
- 核对地址、主机、993 端口、IMAP 服务是否开启、客户端密码/授权码是否有效。
- Outlook/Microsoft 365 或 Gmail 若要求 OAuth2，当前版本无法通过普通密码绕过该限制。
- 企业/学校邮箱可能由管理员禁用第三方客户端连接，需向服务商或管理员核对 IMAP 参数。

**Windows 快捷方式或桌面插件不见了**
- 在开始菜单搜索“邮件助手”；如安装的是旧版，请下载最新 ZIP 并再次运行 `install_windows.ps1`，它会重新创建开始菜单和桌面快捷方式。

**macOS 应用无法打开**
- 确认安装了包含 Tkinter 的 Python；运行 `python3 -c 'import tkinter'` 检查。
- 在项目目录重新运行 `bash install_macos.sh`，然后从 `~/Applications` 打开应用。登录自动启动配置文件位于 `~/Library/LaunchAgents/com.zishengliu.mailassistant.plist`。

## 开发与项目文件

- `mail_briefing.py`：桌面悬浮组件、邮件解析、分类、摘要与正文阅读器。
- `install_windows.ps1`：Windows 安装、登录启动任务及桌面/开始菜单快捷方式。
- `install_macos.sh`：macOS 应用包、依赖、Keychain 支持环境与登录启动配置。
- `launch_widget.bat` / `launch.bat`：Windows 桌面组件与设置窗口启动入口。
- `tests/test_mail_briefing.py`：邮件解析、分类、日期筛选和自定义 IMAP 配置测试。
- `.github/workflows/validate.yml`：Python 语法、单元测试与 macOS 安装脚本语法检查。

欢迎通过 GitHub Issues 提出改进建议。请不要在公开 Issue 中粘贴邮箱授权码、完整邮件正文或个人隐私。
