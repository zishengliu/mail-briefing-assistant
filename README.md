# 邮件晨报助手（Windows 桌面版）

> 独立项目仓库：[`zishengliu/mail-briefing-assistant`](https://github.com/zishengliu/mail-briefing-assistant)

每天早上汇总 Gmail、QQ 邮箱和清华大学学生邮箱中的新邮件，并在 Windows 桌面右下角显示可展开的悬浮小插件。默认状态为紧凑摘要，点击“详情”下拉后可滚动查看发件人、主题、分类与时间；点击邮件可打开对应邮箱。

## 功能

- 支持 Gmail（`imap.gmail.com:993`）、QQ 邮箱（`imap.qq.com:993`）和清华大学学生邮箱（`mails.tsinghua.edu.cn:993`）。
- 默认每天 08:00 检查；可在设置中修改时间。\n- 常驻桌面悬浮卡片：可展开/收起、拖动位置、查看邮件详情、打开邮箱或设置；发现新晨报时自动展开，约 25 秒无操作后收起。
- 分类：紧急/待处理、课程与学术、求职与实习、账单与账户、订阅与通知、个人与其他。
- 首次检查整理最近 24 小时内的邮件；之后按 IMAP UID 追踪新到邮件。
- 使用 IMAP 的只读收件箱和 `BODY.PEEK`，不会主动标记已读、发送、删除或移动邮件。
- 邮件只在本机读取与分类，不调用云端 AI；密码/授权码通过操作系统凭据库保存。GitHub 仓库只保存源代码，不存邮箱数据或凭据。

## Windows 安装

1. 在本仓库点击 **Code → Download ZIP** 下载完整项目并解压，或克隆本仓库。
2. 安装 Python 3.11 或更新版本，并勾选 **Add Python to PATH**。
3. 在该文件夹中运行 `install_windows.ps1`。如果 PowerShell 阻止脚本，可打开 PowerShell 后执行：

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\install_windows.ps1
   ```

4. 安装脚本会将程序复制到 `%LOCALAPPDATA%\MailBriefingAssistant`、创建虚拟环境，并创建 Windows 登录启动任务。
5. 从开始菜单搜索 **邮件晨报助手**，打开设置，填入邮箱地址和对应的客户端密码/授权码，点击保存，再测试连接。
6. 点击“立即检查新邮件”验证。之后只要用户已登录 Windows 且电脑运行，后台会在设定时间检查。电脑关机期间不会联网，登录后若当天尚未生成晨报，会补跑一次。

也可运行 `launch.bat` 打开邮箱设置与晨报列表界面。

### 更新已安装的版本

如果之前安装过旧版系统通知程序，请下载本仓库最新 ZIP 并解压，然后再次运行 `install_windows.ps1`。安装脚本会覆盖本机程序文件并更新登录启动任务；邮箱地址、授权码和本地晨报状态会保留在本机。

## 邮箱设置

### Gmail

轻量版使用 IMAP 应用密码方式，不会让你输入 Google 网页登录密码。应用密码通常要求已开启两步验证，且账号/组织政策允许使用应用密码；若该方式不可用，需要后续接入 Google OAuth 授权。

### QQ 邮箱

在 QQ 邮箱网页版的“设置 → 账户”中开启 IMAP/SMTP 服务并生成授权码。程序需要授权码，不是 QQ 登录密码。

### 清华大学学生邮箱

IMAP 地址为 `mails.tsinghua.edu.cn`，SSL 端口 `993`。请按邮箱网页端的客户端设置/安全设置生成或配置客户端专用密码；具体权限可能受学校账户策略影响。详见[清华大学邮箱官方客户端设置说明](https://mail.tsinghua.edu.cn/coremail/help/clientoption.jsp?locale=zh_CN)。

## 隐私与限制

- 凭据交给本机系统凭据库保存；配置文件不包含密码。
- 状态文件记录每个邮箱上次处理的 UID，并暂存待晨报邮件的必要元数据；`latest_digest.json` 保存最近一次晨报的发件人、主题、分类和时间，方便打开助手后查看。不会保存邮件正文、附件或密码；日志不记录主题、正文或密码。
- 分类目前使用可解释的关键词规则，并非大模型语义分类，可能出现误分类。
- 每次最多处理 100 封新邮件，首次扫描最多检查最近 100 封中的近 24 小时邮件。
- 仅扫描收件箱（INBOX），不读取其他文件夹。
- 桌面小插件依赖 Windows 用户已登录且计划任务运行；退出小插件后，后台任务会结束，下一次登录时会重新启动。
- IMAP 可能被邮箱管理员关闭或要求不同登录方式；若连接失败，请检查客户端专用密码/授权码及邮箱策略。
- Gmail 官方更推荐通过 Google 授权接入；轻量版暂未实现 OAuth，因此个别 Gmail/Workspace 账户可能无法使用 IMAP 应用密码。

## 文件

- `mail_briefing.py`：桌面界面、IMAP 检查、分类、通知和调度。
- `install_windows.ps1`：本地安装与登录启动任务。
- `launch.bat`：打开设置界面。
- `requirements.txt`：Python 依赖。
