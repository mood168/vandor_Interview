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
Dim questionId, companyName, answer, visitDate
questionId = Request.Form("questionId")
companyName = Request.Form("companyName")
answer = Replace(Replace(Request.Form("answer"), "%,元", "元"), "|", "")
visitDate = Request.Form("visitDate")

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

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "'" & Replace(str, "'", "''") & "'"
    End If
End Function

On Error Resume Next

' 先取得廠商ID
Dim vendorId
Set rs = conn.Execute("SELECT VendorID FROM Vendors WHERE VendorName = " & SafeSQL(companyName))

If rs.EOF Then
    Response.Write "{""success"":false,""message"":""找不到對應的廠商資料""}"
    Response.End
End If

vendorId = rs("VendorID")

' 檢查是否已存在訪廠記錄
Dim rsCheck, visitId
Set rsCheck = conn.Execute("SELECT VisitID FROM VisitRecords " & _
                          "WHERE CompanyName = " & SafeSQL(companyName) & " " & _
                          "AND CONVERT(date, VisitDate) = CONVERT(date, " & SafeSQL(visitDate) & ")")

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""檢查訪廠記錄時發生錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

If rsCheck.EOF Then
    ' 新增訪廠主表記錄
    Dim sqlInsertVisit
    sqlInsertVisit = "INSERT INTO VisitRecords (CompanyName, VisitDate, VisitorID, Status, CreatedDate) " & _
                     "VALUES (" & SafeSQL(companyName) & ", " & _
                     SafeSQL(visitDate) & ", " & _
                     Session("UserID") & ", 'Draft', GETDATE()); " & _
                     "SELECT SCOPE_IDENTITY() AS NewID"
    
    Set rsCheck = conn.Execute(sqlInsertVisit)
    
    If Err.Number <> 0 Then
        Response.Write "{""success"":false,""message"":""新增訪廠記錄時發生錯誤: " & Server.HTMLEncode(Err.Description) & """}"
        Response.End
    End If
    
    visitId = rsCheck("NewID")
Else
    visitId = rsCheck("VisitID")
End If

' 檢查是否已有答案
Dim rsAnswer
Set rsAnswer = conn.Execute("SELECT AnswerID FROM VisitAnswers WHERE VisitID = " & visitId & " AND QuestionID = " & questionId)

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""檢查答案時發生錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

' 更新或插入答案
Dim sqlAnswer
If rsAnswer.EOF Then
    ' 插入新答案
    sqlAnswer = "INSERT INTO VisitAnswers (VisitID, QuestionID, Answer, ModifiedDate) " & _
                "VALUES (" & visitId & ", " & questionId & ", " & SafeSQL(answer) & ", GETDATE())"
Else
    ' 更新現有答案
    sqlAnswer = "UPDATE VisitAnswers SET " & _
                "Answer = " & SafeSQL(answer) & ", " & _
                "ModifiedDate = GETDATE() " & _
                "WHERE VisitID = " & visitId & " AND QuestionID = " & questionId
End If

conn.Execute sqlAnswer

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""更新答案時發生錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

' 新增歷史記錄
Dim sqlHistory
sqlHistory = "INSERT INTO VisitAnswerHistory (QuestionID, VendorID, Answer, CreatedBy) " & _
            "VALUES (" & questionId & ", " & vendorId & ", " & SafeSQL(answer) & ", " & Session("UserID") & ")"

conn.Execute sqlHistory

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""新增歷史記錄時發生錯誤: " & Server.HTMLEncode(Err.Description) & """}"
Else
    Response.Write "{""success"":true,""message"":""答案儲存成功"",""visitId"":" & visitId & "}"
End If

' 清理資源
If IsObject(rs) Then
    rs.Close
    Set rs = Nothing
End If

If IsObject(rsCheck) Then
    rsCheck.Close
    Set rsCheck = Nothing
End If

If IsObject(rsAnswer) Then
    rsAnswer.Close
    Set rsAnswer = Nothing
End If

conn.Close
Set conn = Nothing
%> 