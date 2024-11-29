<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

Function HandleError(message)
    Response.Clear
    Response.Write "{""success"": false, ""message"": """ & Replace(message, """", "\""") & """}"
    Response.End
End Function

' 取得表單資料
token = Request.Form("token")
password = Request.Form("password")

If token = "" Then HandleError "無效的重設密碼請求"
If password = "" Then HandleError "請輸入新密碼"

On Error Resume Next

' 開始交易
conn.BeginTrans

' 檢查 token 是否有效
sql = "SELECT UserID FROM PasswordResets " & _
      "WHERE ResetToken = '" & Replace(token, "'", "''") & "' " & _
      "AND ExpiryDate > GETDATE() " & _
      "AND IsUsed = 0"

Set rs = conn.Execute(sql)

If rs.EOF Then
    conn.RollbackTrans
    HandleError "重設密碼連結已失效"
    Response.End
End If

userId = rs("UserID")

' 更新密碼
sql = "UPDATE Users SET " & _
      "Password = '" & Replace(password, "'", "''") & "', " & _
      "ModifiedDate = GETDATE() " & _
      "WHERE UserID = " & userId

conn.Execute sql

If Err.Number <> 0 Then
    conn.RollbackTrans
    HandleError "更新密碼時發生錯誤"
    Response.End
End If

' 標記 token 已使用
sql = "UPDATE PasswordResets SET " & _
      "IsUsed = 1, " & _
      "UsedDate = GETDATE() " & _
      "WHERE ResetToken = '" & Replace(token, "'", "''") & "'"

conn.Execute sql

If Err.Number <> 0 Then
    conn.RollbackTrans
    HandleError "更新 Token 狀態時發生錯誤"
    Response.End
End If

' 提交交易
conn.CommitTrans

' 回傳成功訊息
Response.Write "{""success"": true, ""message"": ""密碼已重設成功""}"

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 