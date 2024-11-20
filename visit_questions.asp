<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得所有分類和問題
Dim rsCategories, rsQuestions
Set rsCategories = conn.Execute("SELECT * FROM QuestionCategories ORDER BY SortOrder")
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>訪廠題庫列表</title>
    <link rel="stylesheet" href="styles/dashboard.css">
    <link rel="stylesheet" href="styles/visit_questions.css">
</head>
<!-- 在 top-bar 區域添加主題切換 -->

<body>
    <div class="dashboard-container">
        <!-- 側邊選單 (與 dashboard.asp 相同) -->
        <!--#include file="aside_menu.asp"-->

        <main class="main-content">
        <header class="top-bar">
            <div class="search-bar">
                <input type="search" placeholder="搜尋...">
            </div>
            <div class="user-actions">
                <!--#include file="theme_switch.asp"-->
                <span class="notification">🔔</span>
                <span class="user-profile">👤</span>
            </div>
        </header>
            <div class="visit-form-container">
                <h1>訪廠記錄表</h1>
                
                <form id="visitForm" action="save_visit.asp" method="post">
                    <div class="company-info">
                        <div class="form-group">
                            <label for="companyName">公司名稱</label>
                            <input type="text" id="companyName" name="companyName" required>
                        </div>
                        <div class="form-group">
                            <label for="visitDate">訪談日期</label>
                            <input type="date" id="visitDate" name="visitDate" required>
                        </div>
                    </div>

                    <div class="questions-container">
                        <%
                        Do While Not rsCategories.EOF
                            Dim categoryID, sql
                            categoryID = rsCategories("CategoryID")
                            sql = "SELECT * FROM VisitQuestions WHERE CategoryID = " & categoryID & " ORDER BY SortOrder"
                            Set rsQuestions = conn.Execute(sql)
                        %>
                            <div class="category-section">
                                <h2><%=rsCategories("CategoryName")%></h2>
                                <% If Not rsCategories("IsRequired") Then %>
                                    <span class="optional-tag">選填</span>
                                <% End If %>

                                <div class="questions">
                                    <% Do While Not rsQuestions.EOF %>
                                        <div class="question-item">
                                            <label>
                                                <%=rsQuestions("QuestionText")%>
                                                <% If rsQuestions("IsRequired") Then %>
                                                    <span class="required">*</span>
                                                <% End If %>
                                            </label>

                                            <% If rsQuestions("HasOptions") Then %>
                                                <select name="q_<%=rsQuestions("QuestionID")%>">
                                                    <option value="">請選擇</option>
                                                    <% 
                                                    Dim options, optionItem
                                                    options = Split(Replace(Replace(rsQuestions("Options"), "[", ""), "]", ""), ",")
                                                    For Each optionItem in options
                                                    %>
                                                        <option value="<%=Replace(Replace(optionItem,"""","")," ","")%>">
                                                            <%=Replace(Replace(optionItem,"""","")," ","")%>
                                                        </option>
                                                    <% Next %>
                                                </select>
                                            <% Else %>
                                                <input type="text" 
                                                       name="q_<%=rsQuestions("QuestionID")%>">
                                            <% End If %>

                                            <% If rsQuestions("CanModify") Then %>
                                                <button type="button" class="edit-btn" 
                                                        onclick="editQuestion(<%=rsQuestions("QuestionID")%>)">
                                                    Save It
                                                </button>
                                            <% End If %>
                                        </div>
                                    <% 
                                        rsQuestions.MoveNext
                                        Loop
                                    %>
                                </div>
                            </div>
                        <%
                            rsCategories.MoveNext
                            Loop
                        %>
                    </div>

                    <div class="form-actions">
                        <button type="submit" class="save-btn">儲存</button>
                        <button type="button" class="cancel-btn" onclick="history.back()">取消</button>
                    </div>
                </form>
            </div>
        </main>
    </div>
</body>
</html> 