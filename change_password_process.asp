<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查是否已登入
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得表單資料
Dim currentPassword, newPassword, confirmPassword
currentPassword = Trim(Request.Form("currentPassword"))
newPassword = Trim(Request.Form("newPassword"))
confirmPassword = Trim(Request.Form("confirmPassword"))

' 基本驗證
If currentPassword = "" Or newPassword = "" Or confirmPassword = "" Then
    Response.Redirect "change_password.asp?error=" & Server.URLEncode("請填寫所有欄位")
    Response.End
End If

' 檢查新密碼是否符合規則
If Not IsPasswordValid(newPassword) Then
    Response.Redirect "change_password.asp?error=" & Server.URLEncode("新密碼不符合密碼規則要求")
    Response.End
End If

' 檢查新密碼確認
If newPassword <> confirmPassword Then
    Response.Redirect "change_password.asp?error=" & Server.URLEncode("新密碼與確認密碼不相符")
    Response.End
End If

' 檢查目前密碼是否正確
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "sp_CheckCurrentPassword"

cmd.Parameters.Append cmd.CreateParameter("@UserID", 3, 1, , Session("UserID"))
cmd.Parameters.Append cmd.CreateParameter("@CurrentPassword", 200, 1, 500, SimpleEncrypt(currentPassword))

Set rs = cmd.Execute

If rs.EOF Or Not rs("IsValid") Then
    Response.Redirect "change_password.asp?error=" & Server.URLEncode("目前密碼不正確")
    Response.End
End If

' 更新密碼
cmd.CommandText = "sp_UpdatePassword"
cmd.Parameters.Append cmd.CreateParameter("@NewPassword", 200, 1, 500, SimpleEncrypt(newPassword))

cmd.Execute

If Err.Number <> 0 Then
    Response.Redirect "change_password.asp?error=" & Server.URLEncode("系統錯誤：" & Err.Description)
    Response.End
End If

' 清理資源
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Set cmd = Nothing

' 清除所有 Session
Session.Abandon

' 重導向到登入頁面
Response.Redirect "login.html?message=" & Server.URLEncode("密碼已成功更新，請重新登入")

Function IsPasswordValid(password)
    ' 檢查密碼規則：至少6位,必須包含大小寫字母和數字
    Dim passwordRegex
    Set passwordRegex = New RegExp
    passwordRegex.Global = True
    passwordRegex.Pattern = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d).{6,}$"
    IsPasswordValid = passwordRegex.Test(password)
    Set passwordRegex = Nothing
End Function

Function SimpleEncrypt(inputText)
    Dim i, charCode
    Dim encryptedText
    encryptedText = ""
    
    For i = 1 To Len(inputText)
        charCode = AscW(Mid(inputText, i, 1))
        charCode = charCode + 3 ' 將字符碼增加1（可以根據需要調整）
        encryptedText = encryptedText & ChrW(charCode)
    Next
    
    SimpleEncrypt = encryptedText
End Function

' 簡單的字符替換解密函數
Function SimpleDecrypt(encryptedText)
    Dim i, charCode
    Dim decryptedText
    decryptedText = ""
    
    For i = 1 To Len(encryptedText)
        charCode = AscW(Mid(encryptedText, i, 1))
        charCode = charCode - 3 ' 將字符碼減少1（必須與加密時相反）
        decryptedText = decryptedText & ChrW(charCode)
    Next
    
    SimpleDecrypt = decryptedText
End Function
%> 