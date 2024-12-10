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
%> 