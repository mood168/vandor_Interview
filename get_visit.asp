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

' 取得訪廠記錄ID
Dim visitId
visitId = Request.QueryString("id")

If visitId = "" Then
    Response.Write "{""success"":false,""message"":""未提供訪廠記錄ID""}"
    Response.End
End If

On Error Resume Next

' 取得訪廠記錄資料
Dim sql
sql = "SELECT vr.*, u.FullName as VisitorName " & _
      "FROM VisitRecords vr " & _
      "LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
      "WHERE vr.VisitID = " & CLng(visitId)

Dim rs
Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

If Not rs.EOF Then
    ' 組織 JSON 回應
    Response.Write "{""success"":true," & _
                   """VisitID"":" & rs("VisitID") & "," & _
                   """CompanyName"":""" & Replace(rs("CompanyName"), """", "\""") & """," & _
                   """VisitorID"":" & rs("VisitorID") & "," & _
                   """VisitorName"":""" & Replace(rs("VisitorName"), """", "\""") & """," & _
                   """Interviewee"":""" & Replace(rs("Interviewee") & "", """", "\""") & """," & _
                   """VisitDate"":""" & FormatDateTime(rs("VisitDate"), 2) & """," & _
                   """Status"":""" & rs("Status") & """}"
Else
    Response.Write "{""success"":false,""message"":""找不到訪廠記錄""}"
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 