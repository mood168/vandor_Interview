<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 防止 XSS 攻擊
Function SanitizeHTML(str)
    If IsNull(str) Or str = "" Then
        SanitizeHTML = ""
        Exit Function
    End If
    
    Dim tmp
    tmp = str
    tmp = Replace(tmp, "<", "&lt;")
    tmp = Replace(tmp, ">", "&gt;")
    tmp = Replace(tmp, """", "&quot;")
    tmp = Replace(tmp, "'", "&#39;")
    tmp = Replace(tmp, "--", "&#45;&#45;")
    tmp = Replace(tmp, ";", "&#59;")
    SanitizeHTML = tmp
End Function

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得訪廠記錄ID
Dim visitId
visitId = Request.QueryString("id")

' 驗證 visitId 是否為數字
If Not IsNumeric(visitId) Then
    Response.Write "無效的訪廠記錄ID"
    Response.End
End If

' 使用參數化查詢取得訪廠記錄
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandText = "SELECT vr.*, u.FullName as VisitorName, " & _
                 "v.UniformNumber, v.Website, v.LogisticsContact, v.MarketingContact, v.ContactPerson " & _
                 "FROM VisitRecords vr " & _
                 "LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
                 "LEFT JOIN Vendors v ON vr.CompanyName = v.VendorName " & _
                 "WHERE vr.VisitID = ?"
cmd.Parameters.Append cmd.CreateParameter("VisitID", adInteger, adParamInput)
cmd.Parameters("VisitID").Value = CLng(visitId)
cmd.Prepared = True

Dim rs
Set rs = cmd.Execute()

If rs.EOF Then
    Response.Write "找不到訪廠記錄"
    Response.End
End If

' 使用參數化查詢取得問題和答案
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandText = "SELECT " & _
                 "qc.CategoryName, " & _
                 "vq.QuestionText, " & _
                 "va.Answer, " & _
                 "va.ModifiedDate " & _
                 "FROM QuestionCategories qc " & _
                 "INNER JOIN VisitQuestions vq ON qc.CategoryID = vq.CategoryID " & _
                 "LEFT JOIN VisitAnswers va ON vq.QuestionID = va.QuestionID " & _
                 "AND va.VisitID = ? " & _
                 "ORDER BY qc.SortOrder, vq.SortOrder"
cmd.Parameters.Append cmd.CreateParameter("VisitID", adInteger, adParamInput)
cmd.Parameters("VisitID").Value = CLng(visitId)
cmd.Prepared = True

Dim rsAnswers
Set rsAnswers = cmd.Execute()
%>

<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline';">
    <title>E-Service 電商訪談表</title>
    <style>
        body {
            font-family: Arial, "Microsoft JhengHei", sans-serif;
            margin: 20px;
            padding: 0;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .container {
            width: 100%;
            max-width: 1200px;
            margin: 0 auto;
        }
        .header {
            text-align: center;
            margin-bottom: 20px;
            width: 100%;
            position: relative;
        }
        .header h1 {
            margin: 0;
            padding: 10px 0;
        }
        .visit-date {
            position: absolute;
            right: 0;
            top: 50%;
            transform: translateY(-50%);
        }
        .company-info {
            margin-bottom: 20px;
            width: 100%;
        }
        .info-row {
            margin-bottom: 10px;
        }
        .info-row label {
            display: inline-block;
            min-width: 120px;
        }
        .section {
            margin-bottom: 30px;
            width: 100%;
        }
        .section h2 {
            margin-bottom: 15px;
            border-bottom: 2px solid #333;
            padding-bottom: 5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        th, td {
            border: 1px solid #333;
            padding: 8px;
            text-align: left;
        }
        th:first-child, td:first-child {
            width: 40%;
        }
        th:last-child, td:last-child {
            width: 60%;
        }
        th {
            background-color: #f0f0f0;
        }
        @media print {
            .no-print {
                display: none;
            }
            body {
                margin: 0;
                padding: 15mm;
            }
        }
        .print-btn {
            padding: 10px 20px;
            background-color: #4a90e2;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin-bottom: 20px;
        }
        .print-btn:hover {
            background-color: #357abd;
        }
    </style>
</head>
<body>
    <button class="print-btn no-print" onclick="window.print()">列印</button>
    
    <div class="container">
        <div class="header">
            <h1>E-Service 電商訪談表</h1>
            <div class="visit-date">
                拜訪日期：<%=Server.HTMLEncode(FormatDateTime(rs("VisitDate"),2))%>
            </div>
        </div>

        <div class="company-info">            
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>電商名稱 / 統編：</label>
                    <span><%=Server.HTMLEncode(rs("CompanyName"))%> / <%=Server.HTMLEncode(rs("UniformNumber"))%></span>
                </div>
                <div style="flex: 1;">
                    <label>網站名稱 / 網址：</label>
                    <span><%=Server.HTMLEncode(rs("Website"))%></span>
                </div>
                <div style="flex: 1;"></div>
            </div>
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>物流聯絡窗口：</label>
                    <span><%=Server.HTMLEncode(rs("LogisticsContact"))%></span>
                </div>
                <div style="flex: 1;">
                    <label>行銷聯絡窗口：</label>
                    <span><%=Server.HTMLEncode(rs("MarketingContact"))%></span>
                </div>
                <div style="flex: 1;"></div>
            </div>
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>客服聯絡窗口：</label>
                    <span><%=Server.HTMLEncode(rs("ContactPerson"))%></span>
                </div>
                <div style="flex: 1;">
                    <label>受訪電商簽名：</label>
                    <span>______________________</span>
                </div>
                <div style="flex: 1;"></div>
            </div>
        </div>

        <%
        Dim currentCategory
        currentCategory = ""
        
        Do While Not rsAnswers.EOF
            If currentCategory <> rsAnswers("CategoryName") Then
                If currentCategory <> "" Then
                    Response.Write "</table></div>"
                End If
                currentCategory = rsAnswers("CategoryName")
        %>
                <div class="section">
                    <h2>【<%=Server.HTMLEncode(currentCategory)%>】</h2>
                    <table>
                        <tr>
                            <th style="width: 40%;">問題</th>
                            <th style="width: 60%;">回答</th>
                        </tr>
        <%
            End If
        %>
                        <tr>
                            <td><%=Server.HTMLEncode(rsAnswers("QuestionText"))%></td>
                            <td><%
                                Dim answer
                                answer = rsAnswers("Answer")
                                If Not IsNull(answer) Then
                                    If InStr(answer, "|") > 0 Then
                                        ' 處理多選答案
                                        Dim answers
                                        answers = Split(answer, "|")
                                        For Each ans in answers
                                            Response.Write Server.HTMLEncode(Replace(ans, ",", " - ")) & "<br>"
                                        Next
                                    Else
                                        Response.Write Server.HTMLEncode(answer)
                                    End If
                                End If
                            %></td>
                        </tr>
        <%
            rsAnswers.MoveNext
        Loop
        
        If currentCategory <> "" Then
            Response.Write "</table></div>"
        End If
        %>
    </div>
</body>
</html>
<%
' 清理資源
Set rs = Nothing
Set rsAnswers = Nothing
Set cmd = Nothing
%> 