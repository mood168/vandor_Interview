<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Write "{""success"":false,""message"":""未登入""}"
    Response.End
End If

' 取得使用者ID
Dim userId
userId = Request.QueryString("id")

If userId = "" Then
    Response.Write "{""success"":false,""message"":""未提供使用者ID""}"
    Response.End
End If

On Error Resume Next

' 取得使用者資料
Dim sql
sql = "SELECT * FROM Users WHERE UserID = " & CLng(userId)

Dim rs
Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

If Not rs.EOF Then
    ' 組織 JSON 回應
    Response.Write "{""success"":true," & _
                   """UserID"":" & rs("UserID") & "," & _
                   """Username"":""" & rs("Username") & """," & _
                   """FullName"":""" & rs("FullName") & """," & _
                   """Phone"":""" & rs("Phone") & """," & _
                   """Email"":""" & rs("Email") & """," & _
                   """Department"":""" & rs("Department") & """," & _
                   """UserRole"":""" & rs("UserRole") & """," & _
                   """IsActive"":" & LCase(rs("IsActive")) & "}"
Else
    Response.Write "{""success"":false,""message"":""找不到使用者資料""}"
End If

' 清理資源
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 