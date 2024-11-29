<%
' SMTP 設定
Const SMTP_SERVER = "smtp.gmail.com"  ' 或其他 SMTP 伺服器
Const SMTP_PORT = 587
Const SMTP_USERNAME = "your-email@gmail.com"  ' 您的 Gmail 帳號
Const SMTP_PASSWORD = "your-app-password"      ' Gmail 應用程式密碼

' 郵件設定
Const MAIL_FROM = "your-email@gmail.com"       ' 寄件者信箱
Const WEBSITE_URL = "http://your-domain.com"   ' 您的網站網址

' 錯誤記錄函數
Sub LogError(message)
    Dim fso, logFile
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile(Server.MapPath("../logs/mail_error.log"), 8, True)
    logFile.WriteLine Now() & " - " & message
    logFile.Close
    Set logFile = Nothing
    Set fso = Nothing
End Sub
%> 