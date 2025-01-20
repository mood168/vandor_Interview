<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得訪廠記錄ID
Dim visitId
visitId = Request.QueryString("id")

' 取得訪廠記錄基本資料
Dim sql
sql = "SELECT vr.*, u.FullName as VisitorName, " & _
      "v.UniformNumber, v.Website, v.LogisticsContact, v.MarketingContact ,v.ContactPerson " & _
      "FROM VisitRecords vr " & _
      "LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
      "LEFT JOIN Vendors v ON vr.CompanyName = v.VendorName " & _
      "WHERE vr.VisitID = " & visitId

Dim rs
Set rs = conn.Execute(sql)

If rs.EOF Then
    Response.Write "找不到訪廠記錄"
    Response.End
End If

' 取得所有問題和答案
Dim sqlAnswers
sqlAnswers = "SELECT " & _
            "qc.CategoryName, " & _
            "vq.QuestionText, " & _
            "va.Answer, " & _
            "va.ModifiedDate " & _
            "FROM QuestionCategories qc " & _
            "INNER JOIN VisitQuestions vq ON qc.CategoryID = vq.CategoryID " & _
            "LEFT JOIN VisitAnswers va ON vq.QuestionID = va.QuestionID " & _
            "AND va.VisitID = " & visitId & " " & _
            "ORDER BY qc.SortOrder, vq.SortOrder"

Dim rsAnswers
Set rsAnswers = conn.Execute(sqlAnswers)
%>

<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>E-Service 訪廠資料表</title>
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
            <h1>E-Service 訪廠資料表</h1>
            <div class="visit-date">
                拜訪日期：<%=FormatDateTime(rs("VisitDate"),2)%>
            </div>
        </div>

        <div class="company-info">            
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>電商名稱 / 統編：</label>
                    <span><%=rs("CompanyName")%> / <%=rs("UniformNumber")%></span>
                </div>
                <div style="flex: 1;">
                    <label>網站名稱 / 網址：</label>
                    <span><%=rs("Website")%></span>
                </div>
                <div style="flex: 1;"></div>
            </div>
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>物流聯絡窗口：</label>
                    <span><%=rs("LogisticsContact")%></span>
                </div>
                <div style="flex: 1;">
                    <label>行銷聯絡窗口：</label>
                    <span><%=rs("MarketingContact")%></span>
                </div>
                <div style="flex: 1;"></div>
            </div>
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>客服聯絡窗口：</label>
                    <span><%=rs("ContactPerson")%></span>
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
                    <h2>【<%=currentCategory%>】</h2>
                    <table>
                        <tr>
                            <th style="width: 40%;">問題</th>
                            <th style="width: 60%;">回答</th>
                        </tr>
        <%
            End If
        %>
                        <tr>
                            <td><%=rsAnswers("QuestionText")%></td>
                            <td><%=rsAnswers("Answer")%></td>
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