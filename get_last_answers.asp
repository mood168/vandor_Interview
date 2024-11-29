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

' 檢查登入狀態
If Session("UserID") = "" Then
    HandleError "請先登入系統"
End If

' 取得公司名稱
companyName = Request.QueryString("companyName")
If companyName = "" Then
    HandleError "未提供公司名稱"
End If

On Error Resume Next

' 查詢最近的答案
sql = "SELECT va.QuestionID, va.Answer, " & _
      "FORMAT(va.ModifiedDate, 'yyyy-MM-dd') AS ModifiedDate " & _
      "FROM VisitAnswers va " & _
      "INNER JOIN VisitRecords vr ON va.VisitID = vr.VisitID " & _
      "WHERE vr.CompanyName = N'" & Replace(companyName, "'", "''") & "' " & _
      "AND va.ModifiedDate = ( " & _
      "    SELECT MAX(va2.ModifiedDate) " & _
      "    FROM VisitAnswers va2 " & _
      "    INNER JOIN VisitRecords vr2 ON va2.VisitID = vr2.VisitID " & _
      "    WHERE vr2.CompanyName = vr.CompanyName " & _
      "    AND va2.QuestionID = va.QuestionID " & _
      ") " & _
      "ORDER BY va.QuestionID"

Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    HandleError "查詢答案時發生錯誤: " & Err.Description
End If

' 組織 JSON 回應
Response.Write "{""success"": true, ""answers"": ["

isFirst = True
Do While Not rs.EOF
    If Not isFirst Then Response.Write ","
    Response.Write "{"
    Response.Write """QuestionID"": " & rs("QuestionID") & ","
    Response.Write """Answer"": """ & Replace(rs("Answer"), """", "\""") & ""","
    Response.Write """ModifiedDate"": """ & rs("ModifiedDate") & """"
    Response.Write "}"
    isFirst = False
    rs.MoveNext
Loop

Response.Write "]}"

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 