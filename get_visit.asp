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

Function FormatDate(dateValue)
    If IsNull(dateValue) Then
        FormatDate = ""
    Else
        FormatDate = Year(dateValue) & "-" & _
                    Right("0" & Month(dateValue), 2) & "-" & _
                    Right("0" & Day(dateValue), 2)
    End If
End Function

' 檢查登入狀態
If Session("UserID") = "" Then
    HandleError "請先登入系統"
End If

' 取得訪廠記錄ID
visitId = Request.QueryString("id")
If visitId = "" Then
    HandleError "未提供訪廠記錄ID"
End If

On Error Resume Next

' SQL 查詢訪廠記錄
sql = "SELECT vr.VisitID, vr.CompanyName, vr.VisitorID, " & _
      "vr.Interviewee, vr.VisitDate, vr.Status " & _
      "FROM VisitRecords vr " & _
      "WHERE vr.VisitID = " & CLng(visitId)

Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    HandleError "查詢資料時發生錯誤: " & Err.Description
End If

If rs.EOF Then
    HandleError "找不到訪廠記錄"
Else
    ' 組織 JSON 回應
    Response.Write "{"
    Response.Write """success"": true,"
    Response.Write """VisitID"": " & rs("VisitID") & ","
    Response.Write """CompanyName"": """ & Replace(rs("CompanyName"), """", "\""") & ""","
    Response.Write """VisitorID"": " & rs("VisitorID") & ","
    Response.Write """Interviewee"": """ & Replace(rs("Interviewee") & "", """", "\""") & ""","
    Response.Write """VisitDate"": """ & FormatDate(rs("VisitDate")) & ""","
    Response.Write """Status"": """ & Replace(rs("Status"), """", "\""") & """"
    Response.Write "}"
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 