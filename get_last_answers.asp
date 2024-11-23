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

' 取得公司名稱
Dim companyName
companyName = Request.QueryString("companyName")

If companyName = "" Then
    Response.Write "{""success"":false,""message"":""未提供公司名稱""}"
    Response.End
End If

On Error Resume Next

' SQL 注入防護
Function SafeSQL(str)
    SafeSQL = "'" & Replace(str, "'", "''") & "'"
End Function

' 查詢最近的答案
Dim sql
sql = "WITH LastAnswers AS (" & _
      "  SELECT " & _
      "    va.QuestionID, " & _
      "    va.Answer, " & _
      "    va.ModifiedDate, " & _
      "    ROW_NUMBER() OVER (PARTITION BY va.QuestionID ORDER BY va.ModifiedDate DESC) as rn " & _
      "  FROM VisitAnswers va " & _
      "  INNER JOIN VisitRecords vr ON va.VisitID = vr.VisitID " & _
      "  WHERE vr.CompanyName = " & SafeSQL(companyName) & _
      ") " & _
      "SELECT QuestionID, Answer, ModifiedDate " & _
      "FROM LastAnswers " & _
      "WHERE rn = 1"

Dim rs
Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

' 組織 JSON 回應
Response.Write "{""success"":true,""answers"":["

Dim isFirst : isFirst = True

Do While Not rs.EOF
    If Not isFirst Then Response.Write ","
    Response.Write "{""QuestionID"":" & rs("QuestionID") & ","
    Response.Write """Answer"":""" & Replace(rs("Answer"), """", "\""") & ""","
    Response.Write """ModifiedDate"":""" & Year(rs("ModifiedDate")) & "-" & _
                                         Right("0" & Month(rs("ModifiedDate")), 2) & "-" & _
                                         Right("0" & Day(rs("ModifiedDate")), 2) & """}"
    isFirst = False
    rs.MoveNext
Loop

Response.Write "]}"

' 清理資源
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 