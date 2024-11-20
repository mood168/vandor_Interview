<%@ Language="VBScript" CodePage="65001" %>
<%
' 清除所有 Session 變數
Session.Abandon

' 重導向到登入頁面
Response.Redirect "login.html"
%> 