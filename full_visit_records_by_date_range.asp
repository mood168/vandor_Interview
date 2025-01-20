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
Dim visitorID
visitorID = Request.QueryString("id")

' 取得該公司所有訪廠記錄
Dim sql
sql = "SELECT vr.*, u.FullName as VisitorName, " & _
      "v.UniformNumber, v.Website, v.LogisticsContact, v.MarketingContact, v.ContactPerson " & _
      "FROM VisitRecords vr " & _
      "LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
      "LEFT JOIN Vendors v ON vr.CompanyName = v.VendorName " & _
      "WHERE vr.VisitID = " & visitorID  & _
      "ORDER BY vr.VisitDate DESC"

Dim rs
Set rs = conn.Execute(sql)


If rs.EOF Then
    Response.Write "找不到訪廠記錄"
    Response.End
End If

' 取得第一筆記錄的基本資料
Dim basicInfo
Set basicInfo = rs
%>

<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>E-Service 完整訪廠記錄表</title>
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
        th {
            background-color: #f0f0f0;
        }
        .visit-record {
            margin-bottom: 30px;
            border: 1px solid #ccc;
            padding: 20px;
            border-radius: 5px;
        }
        .visit-record h3 {
            margin-top: 0;
            border-bottom: 1px solid #ccc;
            padding-bottom: 10px;
        }
        @media print {
            .no-print {
                display: none;
            }
            body {
                margin: 0;
                padding: 15mm;
            }
            .visit-record {
                break-inside: avoid;
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
            <h1>E-Service 完整訪廠記錄表</h1>
        </div>

        <div class="company-info">            
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>電商名稱 / 統編：</label>
                    <span><%=basicInfo("CompanyName")%> / <%=basicInfo("UniformNumber")%></span>
                </div>
                <div style="flex: 1;">
                    <label>網站名稱 / 網址：</label>
                    <span><%=basicInfo("Website")%></span>
                </div>
                <div style="flex: 1;"></div>
            </div>
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>物流聯絡窗口：</label>
                    <span><%=basicInfo("LogisticsContact")%></span>
                </div>
                <div style="flex: 1;">
                    <label>行銷聯絡窗口：</label>
                    <span><%=basicInfo("MarketingContact")%></span>
                </div>
                <div style="flex: 1;"></div>
            </div>
            <div class="info-row" style="display: flex; gap: 20px;">
                <div style="flex: 1;">
                    <label>客服聯絡窗口：</label>
                    <span><%=basicInfo("ContactPerson")%></span>
                </div>
                <div style="flex: 1;"></div>
                <div style="flex: 1;"></div>
            </div>
        </div>

        <div class="section">
            <h2>訪廠記錄列表</h2>
            
            <% 
            ' 重設資料集指標
            rs.MoveFirst
            
            Do While Not rs.EOF 
                ' 取得該次訪廠的問題和答案
                Dim sqlAnswers
                sqlAnswers = "SELECT " & _
                            "qc.CategoryName, " & _
                            "vq.QuestionText, " & _
                            "va.Answer, " & _
                            "va.ModifiedDate " & _
                            "FROM QuestionCategories qc " & _
                            "INNER JOIN VisitQuestions vq ON qc.CategoryID = vq.CategoryID " & _
                            "LEFT JOIN VisitAnswers va ON vq.QuestionID = va.QuestionID " & _
                            "AND va.VisitID = " & rs("VisitID") & " " & _
                            "ORDER BY qc.SortOrder, vq.SortOrder"

                Dim rsAnswers
                Set rsAnswers = conn.Execute(sqlAnswers)
            %>
                <div class="visit-record">
                    <h3>訪廠日期：<%=FormatDateTime(rs("VisitDate"),2)%></h3>
                    <div class="info-row">
                        <label>訪廠人員：</label>
                        <span><%=rs("VisitorName")%></span>
                    </div>
                    <div class="info-row">
                        <label>受訪人：</label>
                        <span><%=rs("Interviewee")%></span>
                    </div>
                    
                    <%
                    Dim currentCategory
                    currentCategory = ""
                    
                    Do While Not rsAnswers.EOF
                        If currentCategory <> rsAnswers("CategoryName") Then
                            If currentCategory <> "" Then
                                Response.Write "</table>"
                            End If
                            currentCategory = rsAnswers("CategoryName")
                    %>
                            <h4><%=currentCategory%></h4>
                            <table>
                                <tr>
                                    <th>問題</th>
                                    <th>回答</th>
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
                        Response.Write "</table>"
                    End If
                    %>
                </div>
            <%
                rs.MoveNext
            Loop
            %>
        </div>
    </div>
</body>
</html>
