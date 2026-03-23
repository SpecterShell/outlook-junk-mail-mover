# Outlook Junk Mail Mover

語言： [English](README.md) | [简体中文](README.zh-CN.md) | 繁體中文

這個腳本會透過 Microsoft Graph 輪詢 Outlook 的 `junkemail` 資料夾，並將符合條件的郵件移回 `inbox`。

- 使用裝置代碼登入方式驗證 Microsoft Graph。
- 檢查 `junkemail` 中最近的郵件。
- 依寄件者地址、寄件者網域、主旨關鍵字或內文關鍵字進行比對。
- 將符合條件的郵件移動到 `inbox`。
- 以定時方式重複執行，或以單次模式執行，方便搭配 cron/systemd 使用。

## 安裝與設定

1. 在 Microsoft Entra / Azure 中註冊一個應用程式，並記下它的 client ID。
2. 在應用程式註冊中完成以下設定：
   - 開啟應用程式的 Overview 頁面，複製 Application (client) ID。
   - 如果你要登入的是 Outlook.com/Hotmail/Live 個人信箱，請確認 Supported account types 包含個人 Microsoft 帳號。
   - 在 Authentication 中啟用 public client flows。
   - 在 API permissions 中加入 Microsoft Graph 的委派權限 `Mail.ReadWrite`。
3. 將 `.env.example` 複製成 `.env`，並填入設定。
   如果你不想手動編輯 `.env`，可以執行設定精靈：

   ```bash
   uv run outlook_junk_mover.py --configure
   ```

## 範例設定

```dotenv
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_TENANT_ID=consumers
OUTLOOK_ALLOWED_SENDERS=john.doe@example.com,noreply@example.com
OUTLOOK_SUBJECT_KEYWORDS=verification code,login code
OUTLOOK_POLL_SECONDS=120
```

對於 Outlook.com/Hotmail/Live 個人信箱，請設定 `OUTLOOK_TENANT_ID=consumers`，並確認應用程式註冊支援個人 Microsoft 帳號。

對於工作或學校信箱，`OUTLOOK_TENANT_ID` 通常應填入實際租用戶 GUID 或已驗證網域，而不是 `common`。`common` 在瀏覽器登入流程中較方便，但在裝置代碼流程中，工作帳號通常需要租用戶專用的 authority。

## 執行方式

執行一次：

```bash
python3 outlook_junk_mover.py --once
```

持續執行：

```bash
python3 outlook_junk_mover.py
```

第一次執行時，Microsoft 會顯示裝置代碼登入提示。完成登入後，腳本會把權杖快取存到 `.tokens/msal_cache.json`。

## 注意事項

- 精確的寄件者地址允許清單比整個網域的允許清單更安全。
- `OUTLOOK_MOVE_ALL=true` 會把最近的所有垃圾郵件都移到收件匣。除非你非常確定，否則通常不建議這麼做。
- 如果你的租用戶透過 Conditional Access 停用了裝置代碼流程，在策略變更之前，這個腳本將無法完成驗證；此時需要調整策略，或改用其他驗證方式。

## 參考

- Microsoft Graph message move API: https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0
- Microsoft Graph permissions reference (`Mail.ReadWrite`): https://learn.microsoft.com/en-us/graph/permissions-reference
- MSAL Python token acquisition: https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens
- Public client / device-code flow guidance: https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-applications
