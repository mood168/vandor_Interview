<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 關閉錯誤頁面
Server.ScriptTimeout = 300
Response.Buffer = True
On Error Resume Next

Function HandleError(message)
    On Error Resume Next
    Response.Clear
    Response.Write "{""success"": false, ""message"": """ & Replace(message, """", "\""") & """}"
    Response.End
End Function

' 檢查登入狀態
If Session("UserID") = "" Then
    HandleError "請先登入系統"
End If

' 取得表單資料
Dim questionId, companyName, answer, visitDate
questionId = Request.Form("questionId")
companyName = Trim(Request.Form("companyName"))
answer = Trim(Request.Form("answer"))
visitDate = Request.Form("visitDate")

' 基本驗證
If questionId = "" Then HandleError "未提供問題ID"
If companyName = "" Then HandleError "未提供公司名稱"
If answer = "" Then HandleError "未提供答案"
If visitDate = "" Then HandleError "未提供訪談日期"

Dim vendorId, visitId

' 先查詢電商ID
Dim sql
sql = "SELECT VendorID FROM Vendors WHERE VendorName = N'" & Replace(companyName, "'", "''") & "' AND IsActive = 1"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn, 1, 1

If Err.Number <> 0 Then
    HandleError "查詢電商資料時發生錯誤: " & Err.Description & " SQL: " & sql
    Response.End
End If

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    HandleError "找不到對應的電商資料：" & companyName
    Response.End
End If

vendorId = rs("VendorID")
rs.Close
Set rs = Nothing

' 開始交易
conn.BeginTrans

' 先查詢或建立訪廠記錄
sql = "SELECT VisitID FROM VisitRecords " & _
      "WHERE CompanyName = N'" & Replace(companyName, "'", "''") & "' " & _
      "AND CONVERT(date, VisitDate) = CONVERT(date, '" & Replace(visitDate, "'", "''") & "')"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn, 1, 1

If Err.Number <> 0 Then
    conn.RollbackTrans
    HandleError "查詢訪廠記錄時發生錯誤: " & Err.Description & " SQL: " & sql
    Response.End
End If

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    
    ' 建立新的訪廠記錄
    sql = "INSERT INTO VisitRecords (CompanyName, VisitDate, VisitorID, Status, CreatedDate) " & _
          "VALUES (N'" & Replace(companyName, "'", "''") & "', " & _
          "CONVERT(datetime, '" & Replace(visitDate, "'", "''") & "'), " & _
          Session("UserID") & ", 'Draft', GETDATE())"
    
    conn.Execute sql
    
    If Err.Number <> 0 Then
        conn.RollbackTrans
        HandleError "建立訪廠記錄時發生錯誤: " & Err.Description & " SQL: " & sql
        Response.End
    End If
    
    ' 取得新建立的記錄ID
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT VisitID FROM VisitRecords " & _
          "WHERE CompanyName = N'" & Replace(companyName, "'", "''") & "' " & _
          "AND CONVERT(date, VisitDate) = CONVERT(date, '" & Replace(visitDate, "'", "''") & "')"
    
    rs.Open sql, conn, 1, 1
    
    If Err.Number <> 0 Or rs.EOF Then
        rs.Close
        Set rs = Nothing
        conn.RollbackTrans
        HandleError "無法取得新建立的訪廠記錄ID"
        Response.End
    End If
    
    visitId = rs("VisitID")
    rs.Close
    Set rs = Nothing
Else
    visitId = rs("VisitID")
    rs.Close
    Set rs = Nothing
End If

' 更新或新增答案
sql = "IF EXISTS (SELECT 1 FROM VisitAnswers WHERE VisitID = " & CStr(visitId) & " AND QuestionID = " & CStr(questionId) & ") " & _
      "UPDATE VisitAnswers SET Answer = N'" & Replace(answer, "'", "''") & "', ModifiedDate = GETDATE() " & _
      "WHERE VisitID = " & CStr(visitId) & " AND QuestionID = " & CStr(questionId) & " " & _
      "ELSE " & _
      "INSERT INTO VisitAnswers (VisitID, QuestionID, Answer, ModifiedDate) " & _
      "VALUES (" & CStr(visitId) & ", " & CStr(questionId) & ", N'" & Replace(answer, "'", "''") & "', GETDATE())"

conn.Execute sql

If Err.Number <> 0 Then
    conn.RollbackTrans
    HandleError "儲存答案時發生錯誤: " & Err.Description & " SQL: " & sql
    Response.End
End If

' 寫入歷史記錄
sql = "INSERT INTO VisitAnswerHistory (QuestionID, VendorID, Answer, CreatedBy, CreatedDate) " & _
      "VALUES (" & CStr(questionId) & ", " & CStr(vendorId) & ", " & _
      "N'" & Replace(answer, "'", "''") & "', " & Session("UserID") & ", GETDATE())"

conn.Execute sql

If Err.Number <> 0 Then
    conn.RollbackTrans
    HandleError "寫入歷史記錄時發生錯誤: " & Err.Description & " SQL: " & sql
    Response.End
End If

' 提交交易
conn.CommitTrans

' 回傳成功訊息
Response.Clear
Response.Write "{""success"": true, ""message"": ""答案已儲存""}"

conn.Close
Set conn = Nothing
%> 