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
Dim companyName, questionId, answer, visitDate
companyName = Request.Form("companyName")
questionId = Request.Form("questionId")
answer = Request.Form("answer")
visitDate = Request.Form("visitDate")

If companyName = "" Or questionId = "" Or answer = "" Or visitDate = "" Then
    Response.Write "{""success"":false,""message"":""缺少必要參數""}"
    Response.End
End If

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "'" & Replace(str, "'", "''") & "'"
    End If
End Function

On Error Resume Next

' 先檢查是否已有訪廠記錄
Dim sql, rs, visitId
sql = "SELECT VisitID FROM VisitRecords " & _
      "WHERE CompanyName = " & SafeSQL(companyName) & " " & _
      "AND CONVERT(date, VisitDate) = CONVERT(date, " & SafeSQL(visitDate) & ")"

Set rs = conn.Execute(sql)

If rs.EOF Then
    ' 建立新的訪廠記錄
    sql = "INSERT INTO VisitRecords (CompanyName, VisitDate, VisitorID, Status) " & _
          "VALUES (" & SafeSQL(companyName) & ", " & SafeSQL(visitDate) & ", " & _
          Session("UserID") & ", 'Draft'); SELECT SCOPE_IDENTITY() AS NewID"
    
    Set rs = conn.Execute(sql)
    visitId = rs("NewID")
Else
    visitId = rs("VisitID")
End If

' 更新答案
sql = "MERGE VisitAnswers AS target " & _
      "USING (SELECT " & visitId & " AS VisitID, " & questionId & " AS QuestionID) AS source " & _
      "ON target.VisitID = source.VisitID AND target.QuestionID = source.QuestionID " & _
      "WHEN MATCHED THEN " & _
      "    UPDATE SET Answer = " & SafeSQL(answer) & ", ModifiedDate = GETDATE() " & _
      "WHEN NOT MATCHED THEN " & _
      "    INSERT (VisitID, QuestionID, Answer) " & _
      "    VALUES (" & visitId & ", " & questionId & ", " & SafeSQL(answer) & ");"

conn.Execute sql

' 更新歷史記錄
sql = "INSERT INTO VisitAnswerHistory (QuestionID, VendorID, Answer, CreatedBy) " & _
      "SELECT " & questionId & ", v.VendorID, " & SafeSQL(answer) & ", " & Session("UserID") & " " & _
      "FROM Vendors v WHERE v.VendorName = " & SafeSQL(companyName)

conn.Execute sql

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
Else
    Response.Write "{""success"":true,""message"":""答案儲存成功"",""visitId"":" & visitId & "}"
End If

If IsObject(rs) Then
    rs.Close
    Set rs = Nothing
End If

conn.Close
Set conn = Nothing
%> 