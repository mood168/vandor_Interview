<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<meta charset="UTF-8">
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Write "{""success"":false,""message"":""未登入""}"
    Response.End
End If

' 取得表單資料
Dim userId, username, password, fullName, phone, email, department, userRole

userId = Request.Form("userId")
username = Request.Form("username")
password = Request.Form("password")
fullName = Request.Form("fullName")
phone = Request.Form("phone")
email = Request.Form("email")
department = Request.Form("department")
userRole = Request.Form("userRole")
userStatus = Request.Form("userStatus")

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "N'" & Replace(str, "'", "''") & "'"
    End If
End Function

Dim sql
If userId = "" Then
    ' 新增使用者
    sql = "INSERT INTO Users (Username, Password, FullName, Phone, Email, Department, UserRole, IsActive) VALUES (" & _
          SafeSQL(username) & ", " & _
          SafeSQL(password) & ", " & _
          SafeSQL(fullName) & ", " & _
          SafeSQL(phone) & ", " & _
          SafeSQL(email) & ", " & _
          SafeSQL(department) & ", " & _
          SafeSQL(userRole) & ", 1)"
Else
    ' 更新使用者
    sql = "UPDATE Users SET " & _
          "Username = " & SafeSQL(username) & ", " & _
          "Password = " & SafeSQL(password) & ", " & _
          "FullName = " & SafeSQL(fullName) & ", " & _
          "Phone = " & SafeSQL(phone) & ", " & _
          "Email = " & SafeSQL(email) & ", " & _
          "Department = " & SafeSQL(department) & ", " & _
          "UserRole = " & SafeSQL(userRole) & ", " & _
          "IsActive = " & SafeSQL(userStatus) & ", " & _
          "ModifiedDate = GETDATE() " & _
          "WHERE UserID = " & CLng(userId)
End If

' 執行 SQL
conn.Execute sql

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
Else
    Response.Write "{""success"":true,""message"":""資料儲存成功""}"
    Response.Redirect "user_management.asp"
End If

conn.Close
Set conn = Nothing
%> 