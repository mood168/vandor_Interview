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

' 取得訪廠記錄ID或公司名稱
Dim visitId, companyName
visitId = Request.QueryString("id")
companyName = Request.QueryString("companyName")

If visitId = "" And companyName = "" Then
    Response.Write "{""success"":false,""message"":""未提供訪廠記錄ID或公司名稱""}"
    Response.End
End If

On Error Resume Next

Dim sql
If visitId <> "" Then
    ' 根據訪廠記錄ID取得資料
    sql = "SELECT vr.*, u.FullName as VisitorName, " & _
          "(SELECT TOP 1 va.Answer " & _
          "FROM VisitAnswers va " & _
          "WHERE va.VisitID = vr.VisitID " & _
          "ORDER BY va.ModifiedDate DESC) as LastAnswer " & _
          "FROM VisitRecords vr " & _
          "LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
          "WHERE vr.VisitID = " & CLng(visitId)
Else
    ' 根據公司名稱取得最近的答案
    sql = "SELECT q.QuestionID, q.QuestionText, q.CategoryID, " & _
          "vah.Answer, vah.CreatedDate as ModifiedDate " & _
          "FROM VisitAnswerHistory vah " & _
          "INNER JOIN VisitQuestions q ON vah.QuestionID = q.QuestionID " & _
          "INNER JOIN Vendors v ON vah.VendorID = v.VendorID " & _
          "WHERE v.VendorName = '" & Replace(companyName, "'", "''") & "' " & _
          "AND vah.CreatedDate = ( " & _
          "    SELECT MAX(CreatedDate) " & _
          "    FROM VisitAnswerHistory " & _
          "    WHERE QuestionID = vah.QuestionID " & _
          "    AND VendorID = vah.VendorID " & _
          ") " & _
          "ORDER BY q.CategoryID, q.SortOrder"
End If

Dim rs
Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

If Not rs.EOF Then
    If visitId <> "" Then
        ' 回傳單筆訪廠記錄
        Response.Write "{""success"":true," & _
                       """VisitID"":" & rs("VisitID") & "," & _
                       """CompanyName"":""" & Replace(rs("CompanyName"), """", "\""") & """," & _
                       """VisitorID"":" & rs("VisitorID") & "," & _
                       """VisitorName"":""" & Replace(rs("VisitorName"), """", "\""") & """," & _
                       """VisitDate"":""" & FormatDateTime(rs("VisitDate"), 2) & """," & _
                       """Status"":""" & rs("Status") & """," & _
                       """LastAnswer"":""" & Replace(rs("LastAnswer") & "", """", "\""") & """}"
    Else
        ' 回傳公司的所有最近答案
        Response.Write "{""success"":true,""answers"":["
        Dim first : first = True
        Do While Not rs.EOF
            If Not first Then Response.Write ","
            Response.Write "{""QuestionID"":" & rs("QuestionID") & "," & _
                          """QuestionText"":""" & Replace(rs("QuestionText"), """", "\""") & """," & _
                          """Answer"":""" & Replace(rs("Answer"), """", "\""") & """," & _
                          """ModifiedDate"":""" & FormatDateTime(rs("ModifiedDate"), 2) & """}"
            first = False
            rs.MoveNext
        Loop
        Response.Write "]}"
    End If
Else
    If visitId <> "" Then
        Response.Write "{""success"":false,""message"":""找不到訪廠記錄""}"
    Else
        Response.Write "{""success"":true,""answers"":[]}"
    End If
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 