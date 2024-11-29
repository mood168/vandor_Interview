<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<!--#include file="2D34D3E4/mail_config.asp"-->
<%
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

Function HandleError(message)
    Response.Clear
    Response.Write "{""success"": false, ""message"": """ & Replace(message, """", "\""") & """}"
    Response.End
End Function

Function GenerateResetToken()
    Dim token, i
    Const chars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Randomize
    token = ""
    For i = 1 To 32
        token = token & Mid(chars, Int(Rnd * Len(chars)) + 1, 1)
    Next
    GenerateResetToken = token
End Function

' 取得表單資料
Dim username, email
username = Trim(Request.Form("username"))
email = Trim(Request.Form("email"))

' 驗證輸入
If username = "" Then HandleError "請輸入使用者名稱"
If email = "" Then HandleError "請輸入電子郵件"

On Error Resume Next

' 檢查使用者是否存在
sql = "SELECT UserID, Email, FullName FROM Users " & _
      "WHERE Username = '" & Replace(username, "'", "''") & "' " & _
      "AND Email = '" & Replace(email, "'", "''") & "' " & _
      "AND IsActive = 1"

Set rs = conn.Execute(sql)

If rs.EOF Then
    HandleError "找不到符合的使用者資料"
    Response.End
End If

' 生成重設密碼的 Token
Dim resetToken, userId, userEmail, fullName
resetToken = GenerateResetToken()
userId = rs("UserID")
userEmail = rs("Email")
fullName = rs("FullName")

' 儲存重設密碼的 Token
sql = "INSERT INTO PasswordResets (UserID, ResetToken, ExpiryDate) " & _
      "VALUES (" & userId & ", '" & resetToken & "', DATEADD(hour, 24, GETDATE()))"

conn.Execute sql

If Err.Number <> 0 Then
    HandleError "儲存重設密碼 Token 時發生錯誤"
    Response.End
End If

' 發送重設密碼郵件
Dim objMail
Set objMail = Server.CreateObject("CDO.Message")

' 設定 SMTP 伺服器
With objMail.Configuration.Fields
    ' 使用 Gmail SMTP
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTP_USERNAME
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTP_PASSWORD
    .Update
End With

' 設定郵件內容
objMail.From = MAIL_FROM
objMail.To = userEmail
objMail.Subject = "重設密碼通知"
objMail.HTMLBody = "親愛的 " & fullName & "：<br><br>" & _
                  "您已要求重設密碼。請點擊以下連結重設您的密碼：<br><br>" & _
                  "<a href='" & WEBSITE_URL & "/reset_password.asp?token=" & resetToken & "'>" & _
                  "重設密碼</a><br><br>" & _
                  "此連結將在24小時後失效。如果您沒有要求重設密碼，請忽略此郵件。<br><br>" & _
                  "此為系統自動發送的郵件，請勿直接回覆。"

' 嘗試發送郵件
On Error Resume Next
objMail.Send

If Err.Number <> 0 Then
    ' 記錄郵件發送錯誤
    LogError "發送重設密碼郵件失敗: " & Err.Description
    HandleError "發送重設密碼郵件失敗，請聯繫系統管理員"
    Response.End
End If

Set objMail = Nothing

' 回傳成功訊息
Response.Write "{""success"": true, ""message"": ""重設密碼連結已寄送至您的信箱""}"

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 