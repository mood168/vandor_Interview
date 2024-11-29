<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"
Response.Buffer = True  ' 啟用緩衝

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

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "N'" & Replace(str, "'", "''") & "'"
    End If
End Function

' 日期格式化
Function FormatSQLDate(dateStr)
    If IsDate(dateStr) Then
        Dim d
        d = CDate(dateStr)
        FormatSQLDate = "CONVERT(datetime, '" & _
                       Year(d) & "-" & _
                       Right("0" & Month(d), 2) & "-" & _
                       Right("0" & Day(d), 2) & _
                       "', 120)"
    Else
        FormatSQLDate = "NULL"
    End If
End Function

On Error Resume Next

' 檢查登入狀態
If Session("UserID") = "" Then
    HandleError "未登入"
    Response.End
End If

' 取得表單資料
Dim visitId, companyName, visitorId, interviewee, visitDate, statusMsg
visitId = Request.Form("editVisitId")
companyName = Trim(Request.Form("editCompanyName"))
visitorId = Trim(Request.Form("editVisitorId"))
interviewee = Trim(Request.Form("editInterviewee"))
visitDate = Trim(Request.Form("editVisitDate"))
statusMsg = Trim(Request.Form("editStatus"))

LogDebug("Received Data - visitId: " & visitId & ", companyName: " & companyName & _
         ", visitorId: " & visitorId & ", visitDate: " & visitDate & ", statusMsg: " & statusMsg)

' 驗證必要參數
If visitId = "" Then
    ' 如果是新增模式，檢查公司名稱
    If companyName = "" Then 
        HandleError "請輸入公司名稱"
        Response.End
    End If
End If

If visitorId = "" Then 
    HandleError "請選擇訪廠人員"
    Response.End
End If
If visitDate = "" Then 
    HandleError "請選擇訪廠日期"
    Response.End
End If
If statusMsg = "" Then 
    HandleError "請選擇狀態"
    Response.End
End If

' 檢查 visitorId 是否存在於 Users 表格中
Dim rs
Set rs = conn.Execute("SELECT UserID FROM Users WHERE UserID = " & CLng(visitorId))
If rs.EOF Then
    HandleError "選擇的訪廠人員不存在"
    Response.End
End If
rs.Close

' 如果是編輯模式，先取得原有的公司名稱
If visitId <> "" Then
    Set rs = conn.Execute("SELECT CompanyName FROM VisitRecords WHERE VisitID = " & CLng(visitId))
    If Not rs.EOF Then
        companyName = rs("CompanyName")
        LogDebug("Found existing CompanyName = " & companyName)
    End If
    rs.Close
End If

' 開始交易
conn.BeginTrans

' 更新或新增訪廠記錄
Dim sql
If visitId <> "" Then
    ' 更新現有記錄
    sql = "UPDATE VisitRecords SET " & _
          "VisitorID = " & CLng(visitorId) & ", " & _
          "Interviewee = " & SafeSQL(interviewee) & ", " & _
          "VisitDate = " & FormatSQLDate(visitDate) & ", " & _
          "Status = " & SafeSQL(statusMsg) & ", " & _
          "ModifiedDate = GETDATE() " & _
          "WHERE VisitID = " & CLng(visitId)
Else
    ' 新增記錄
    sql = "INSERT INTO VisitRecords (CompanyName, VisitorID, Interviewee, VisitDate, Status, CreatedDate) " & _
          "VALUES (" & SafeSQL(companyName) & ", " & CLng(visitorId) & ", " & _
          SafeSQL(interviewee) & ", " & FormatSQLDate(visitDate) & ", " & _
          SafeSQL(statusMsg) & ", GETDATE())"
End If

LogDebug("SQL = " & sql)

conn.Execute(sql)

' If Err.Number <> 0 Then
'     conn.RollbackTrans
'     HandleError "資料庫錯誤: " & Server.HTMLEncode(Err.Description) & " SQL: " & sql
'     Response.End
' End If

' 提交交易
conn.CommitTrans

' 確保在這之前沒有其他輸出
Response.Clear
' 回傳成功訊息
Response.Write "{""success"": true, ""message"": ""訪廠記錄儲存成功""}"
Response.End

Set rs = Nothing
conn.Close
Set conn = Nothing
%> 