<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%

Response.CharSet = "utf-8"

' 確保是 POST 請求
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得表單資料
Dim username, password
username = Trim(Request.Form("username"))
password = Trim(Request.Form("password"))

' 基本驗證
If username = "" Or password = "" Then
    Response.Redirect "login.html?error=" & Server.URLEncode("請輸入帳號和密碼")
    Response.End
End If

' 取得客戶端 IP
Dim userIP
userIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userIP = "" Then
    userIP = Request.ServerVariables("REMOTE_ADDR")
End If

On Error Resume Next

' 建立資料庫連線
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "sp_UserLogin"

' 設定參數
cmd.Parameters.Append cmd.CreateParameter("@Username", 200, 1, 500, username)
cmd.Parameters.Append cmd.CreateParameter("@Password", 200, 1, 500, password)
cmd.Parameters.Append cmd.CreateParameter("@LoginIP", 200, 1, 50, userIP)

' 執行預存程序
Set rs = cmd.Execute

If Err.Number <> 0 Then
    ' 資料庫錯誤處理
    Response.Redirect "login.html?error=" & Server.URLEncode("系統錯誤：" & Err.Description)
    Response.End
End If

' Response.Write "Request Method: " & Request.ServerVariables("REQUEST_METHOD") & "<br>"
' Response.Write "Username: " & Request.Form("username") & "<br>"
' Response.Write "Password: " & Request.Form("password") & "<br>"
' Response.End

On Error Goto 0

' 檢查登入結果
If Not rs.EOF Then
    If rs("LoginStatus") Then
        ' �查密碼是否過期
        If IsPasswordExpired(rs("LastPasswordChangeDate")) Then
            Response.Redirect "change_password.asp?expired=1"
            Response.End
        End If
        
        ' 登入成功，設定 Session
        Session("UserID") = rs("UserID")
        Session("Username") = rs("Username")
        Session("FullName") = rs("FullName")
        Session("UserRole") = rs("UserRole")
        Session("LoginTime") = Now()
        
        ' 清理資源
        rs.Close
        Set rs = Nothing
        Set cmd = Nothing
        
        ' 重導向到首頁
        Response.Redirect "dashboard.asp"
    Else
        ' 登入失敗
        Response.Redirect "login.html?error=" & Server.URLEncode(rs("LoginMessage"))
    End If
Else
    Response.Redirect "login.html?error=" & Server.URLEncode("登入驗證失敗")
End If

' 清理資源
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Set cmd = Nothing


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

' 檢查密碼是否過期(超過3個月)
Function IsPasswordExpired(lastChangeDate)
    If IsNull(lastChangeDate) Then
        IsPasswordExpired = True
        Exit Function
    End If
    
    Dim expiryDays: expiryDays = 90 ' 3個月 = 90天
    IsPasswordExpired = DateDiff("d", lastChangeDate, Now()) > expiryDays
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