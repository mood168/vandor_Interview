<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Write "{""success"":false,""message"":""未登入""}"
    Response.End
End If

Dim vendorId
vendorId = Request.Form("vendorId")

On Error Resume Next

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "sp_ActivateVendor"

cmd.Parameters.Append cmd.CreateParameter("@VendorID", 3, 1, , vendorId)
cmd.Parameters.Append cmd.CreateParameter("@ModifiedBy", 3, 1, , Session("UserID"))

cmd.Execute

If Err.Number = 0 Then
    Response.Write "{""success"":true}"
Else
    Response.Write "{""success"":false,""message"":""" & Server.HTMLEncode(Err.Description) & """}"
End If

conn.Close
Set conn = Nothing
%>
