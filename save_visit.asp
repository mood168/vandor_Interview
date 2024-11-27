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

' 取得表單資料
Dim visitId, companyName, visitorId, interviewee, visitDate, status
visitId = Request.Form("visitId")
companyName = Request.Form("companyName")
visitorId = Request.Form("visitorId")
interviewee = Request.Form("interviewee")
visitDate = Request.Form("visitDate")
status = Request.Form("status")

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "'" & Replace(str, "'", "''") & "'"
    End If
End Function

On Error Resume Next

' 更新訪廠記錄
Dim sql
sql = "UPDATE VisitRecords SET " & _
      "CompanyName = " & SafeSQL(companyName) & ", " & _
      "VisitorID = " & SafeSQL(visitorId) & ", " & _
      "Interviewee = " & SafeSQL(interviewee) & ", " & _
      "VisitDate = " & SafeSQL(visitDate) & ", " & _
      "Status = " & SafeSQL(status) & ", " & _
      "ModifiedDate = GETDATE() " & _
      "WHERE VisitID = " & CLng(visitId)

conn.Execute sql

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
Else
    Response.Write "{""success"":true,""message"":""訪廠記錄更新成功""}"
End If

conn.Close
Set conn = Nothing
%> 