# Outlook Junk Mail Mover

语言： [English](README.md) | 简体中文 | [繁體中文](README.zh-TW.md)

这个脚本会通过 Microsoft Graph 轮询 Outlook 的 `junkemail` 文件夹，并将匹配的邮件移回 `inbox`。

- 使用设备代码登录方式认证 Microsoft Graph。
- 检查 `junkemail` 中的最近邮件。
- 按发件人地址、发件人域名、主题关键字或正文关键字进行匹配。
- 将匹配到的邮件移动到 `inbox`。
- 按定时周期重复运行，或者以单次模式运行，方便配合 cron/systemd 使用。

## 安装与配置

1. 在 Microsoft Entra / Azure 中注册一个应用，并记下它的客户端 ID。
2. 在应用注册中完成以下设置：
   - 打开应用的 Overview 页面，复制 Application (client) ID。
   - 如果你要登录的是 Outlook.com/Hotmail/Live 个人邮箱，确保 Supported account types 包含个人 Microsoft 账号。
   - 在 Authentication 中启用 public client flows。
   - 在 API permissions 中添加 Microsoft Graph 的委托权限 `Mail.ReadWrite`。
3. 将 `.env.example` 复制为 `.env`，并填写配置。
   如果你不想手动编辑 `.env`，可以运行配置向导：

   ```bash
   uv run outlook_junk_mover.py --configure
   ```

## 示例配置

```dotenv
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_TENANT_ID=consumers
OUTLOOK_ALLOWED_SENDERS=john.doe@example.com,noreply@example.com
OUTLOOK_SUBJECT_KEYWORDS=verification code,login code
OUTLOOK_POLL_SECONDS=120
```

对于 Outlook.com/Hotmail/Live 个人邮箱，请设置 `OUTLOOK_TENANT_ID=consumers`，并确保应用注册支持个人 Microsoft 账号。

对于工作或学校邮箱，`OUTLOOK_TENANT_ID` 一般应填写真实租户 GUID 或已验证域名，而不是 `common`。`common` 在浏览器登录流程中较方便，但在设备代码流程中，工作账户通常需要租户专用的 authority。

## 运行

单次运行：

```bash
python3 outlook_junk_mover.py --once
```

持续运行：

```bash
python3 outlook_junk_mover.py
```

第一次运行时，Microsoft 会显示设备代码登录提示。完成登录后，脚本会把令牌缓存保存到 `.tokens/msal_cache.json`。

## 注意事项

- 精确的发件人地址允许名单比整域名允许名单更安全。
- `OUTLOOK_MOVE_ALL=true` 会把最近的所有垃圾邮件都移动到收件箱。除非你非常确定，否则通常不建议这么做。
- 如果你的租户通过 Conditional Access 禁用了设备代码流程，在策略修改之前，这个脚本将无法完成认证；此时需要调整策略或改用其他认证方式。

## 参考

- Microsoft Graph message move API: https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0
- Microsoft Graph permissions reference (`Mail.ReadWrite`): https://learn.microsoft.com/en-us/graph/permissions-reference
- MSAL Python token acquisition: https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens
- Public client / device-code flow guidance: https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-applications
