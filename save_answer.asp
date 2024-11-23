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
Dim questionId, companyName, answer
questionId = Request.Form("questionId")
companyName = Request.Form("companyName")
answer = Request.Form("answer")

' 基本驗證
If questionId = "" Then
    Response.Write "{""success"":false,""message"":""未提供問題ID""}"
    Response.End
End If

If companyName = "" Then
    Response.Write "{""success"":false,""message"":""未提供公司名稱""}"
    Response.End
End If

If answer = "" Then
    Response.Write "{""success"":false,""message"":""答案不能為空""}"
    Response.End
End If

On Error Resume Next

' SQL 注入防護
Function SafeSQL(str)
    SafeSQL = "'" & Replace(str, "'", "''") & "'"
End Function

' 檢查是否已存在訪廠記錄
Dim rsCheck, visitId
Set rsCheck = conn.Execute("SELECT VisitID FROM VisitRecords WHERE CompanyName = " & SafeSQL(companyName))

If rsCheck.EOF Then
    ' 新增訪廠主表記錄
    Dim sqlInsertVisit
    sqlInsertVisit = "INSERT INTO VisitRecords (CompanyName, VisitDate, VisitorID, Status, CreatedDate) " & _
                     "VALUES (" & SafeSQL(companyName) & ", GETDATE(), " & Session("UserID") & ", 'Draft', GETDATE())"
    
    conn.Execute sqlInsertVisit
    
    ' 取得新插入的 ID
    Set rsCheck = conn.Execute("SELECT MAX(VisitID) AS NewID FROM VisitRecords WHERE CompanyName = " & SafeSQL(companyName))
    visitId = rsCheck("NewID")
Else
    visitId = rsCheck("VisitID")
End If

' 檢查是否已有答案
Dim rsAnswer
Set rsAnswer = conn.Execute("SELECT AnswerID FROM VisitAnswers WHERE VisitID = " & visitId & " AND QuestionID = " & questionId)

If rsAnswer.EOF Then
    ' 插入新答案
    Dim sqlInsert
    sqlInsert = "INSERT INTO VisitAnswers (VisitID, QuestionID, Answer, ModifiedDate) " & _
                "VALUES (" & visitId & ", " & questionId & ", " & SafeSQL(answer) & ", GETDATE())"
    conn.Execute sqlInsert
Else
    ' 更新現有答案
    Dim sqlUpdate
    sqlUpdate = "UPDATE VisitAnswers SET " & _
                "Answer = " & SafeSQL(answer) & ", " & _
                "ModifiedDate = GETDATE() " & _
                "WHERE VisitID = " & visitId & " AND QuestionID = " & questionId
    conn.Execute sqlUpdate
End If

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
Else
    Response.Write "{""success"":true,""message"":""答案儲存成功""}"
End If

conn.Close
Set conn = Nothing
%> 