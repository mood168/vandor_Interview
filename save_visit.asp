<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

Function HandleError(message)
    Response.Clear
    Response.Write "{""success"": false, ""message"": """ & Replace(message, """", "\""") & """}"
    Response.End
End Function

Function LogDebug(message)
    ' 將除錯訊息寫入伺服器日誌檔，而不是輸出到回應中
    Dim fso, debugFile
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set debugFile = fso.OpenTextFile(Server.MapPath("debug.log"), 8, True)
    debugFile.WriteLine(Now() & " - " & message)
    debugFile.Close
    Set debugFile = Nothing
    Set fso = Nothing
End Function

' 檢查登入狀態
If Session("UserID") = "" Then
    HandleError "未登入"
End If

' 取得表單資料
Dim visitId, companyName, visitorId, interviewee, visitDate, status
visitId = Request.Form("editVisitId")
companyName = Trim(Request.Form("editCompanyName"))
visitorId = Trim(Request.Form("editVisitorId"))
interviewee = Trim(Request.Form("editInterviewee"))
visitDate = Trim(Request.Form("editVisitDate"))
status = Trim(Request.Form("editStatus"))

' 驗證必要參數
If visitId = "" Then
    ' 如果是新增模式，檢查公司名稱
    If companyName = "" Then 
        HandleError "請輸入公司名稱"
    End If
End If

If visitorId = "" Then HandleError "請選擇訪廠人員"
If visitDate = "" Then HandleError "請選擇訪廠日期"
If status = "" Then HandleError "請選擇狀態"

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "N'" & Replace(str, "'", "''") & "'"
    End If
End Function

On Error Resume Next

' 如果是編輯模式，先取得原有的公司名稱
If visitId <> "" Then
    Dim rs
    Set rs = conn.Execute("SELECT CompanyName FROM VisitRecords WHERE VisitID = " & CLng(visitId))
    If Not rs.EOF Then
        companyName = rs("CompanyName")
        LogDebug("Found existing CompanyName = " & companyName)
    End If
    rs.Close
    Set rs = Nothing
End If

' 更新或新增訪廠記錄
Dim sql
If visitId <> "" Then
    ' 更新現有記錄
    sql = "UPDATE VisitRecords SET " & _
          "VisitorID = " & visitorId & ", " & _
          "Interviewee = " & SafeSQL(interviewee) & ", " & _
          "VisitDate = " & SafeSQL(visitDate) & ", " & _
          "Status = " & SafeSQL(status) & ", " & _
          "ModifiedDate = GETDATE() " & _
          "WHERE VisitID = " & CLng(visitId)
Else
    ' 新增記錄
    sql = "INSERT INTO VisitRecords (CompanyName, VisitorID, Interviewee, VisitDate, Status, CreatedDate) " & _
          "VALUES (" & SafeSQL(companyName) & ", " & visitorId & ", " & _
          SafeSQL(interviewee) & ", " & SafeSQL(visitDate) & ", " & SafeSQL(status) & ", GETDATE())"
End If

LogDebug("SQL = " & sql)

conn.Execute sql

If Err.Number <> 0 Then
    HandleError "資料庫錯誤: " & Server.HTMLEncode(Err.Description) & " SQL: " & sql
    Response.End
End If

' 回傳成功訊息
Response.Clear
Response.Write "{""success"": true, ""message"": ""訪廠記錄儲存成功""}"

conn.Close
Set conn = Nothing
%> 