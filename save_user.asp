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
    sql = "INSERT INTO Users (Username, Password, FullName, Phone, Email, Department, UserRole, IsActive, LastPasswordChangeDate) VALUES (" & _
          SafeSQL(SimpleEncrypt(username)) & ", " & _
          SafeSQL(SimpleEncrypt(password)) & ", " & _
          SafeSQL(fullName) & ", " & _
          SafeSQL(phone) & ", " & _
          SafeSQL(email) & ", " & _
          SafeSQL(department) & ", " & _
          SafeSQL(userRole) & ", 1, " & _
          "GETDATE())"
Else
    ' 更新使用者
    sql = "UPDATE Users SET " & _
          "Username = " & SafeSQL(SimpleEncrypt(username)) & ", " & _
          "Password = " & SafeSQL(SimpleEncrypt(password)) & ", " & _
          "FullName = " & SafeSQL(fullName) & ", " & _
          "Phone = " & SafeSQL(phone) & ", " & _
          "Email = " & SafeSQL(email) & ", " & _
          "Department = " & SafeSQL(department) & ", " & _
          "UserRole = " & SafeSQL(userRole) & ", " & _
          "IsActive = " & SafeSQL(userStatus) & ", " & _
          "LastPasswordChangeDate = GETDATE(), " & _
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

Function IsPasswordValid(password)
    ' 檢查密碼規則：至少6位,必須包含大小寫字母和數字
    Dim passwordRegex
    Set passwordRegex = New RegExp
    passwordRegex.Global = True
    passwordRegex.Pattern = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d).{6,}$"
    IsPasswordValid = passwordRegex.Test(password)
    Set passwordRegex = Nothing
End Function

Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "%,/,\,(,),<,>,',--,^,&,?,;,:," & Chr(34) & "," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function
%> 