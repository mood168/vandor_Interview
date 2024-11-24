<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

On Error Resume Next

' 取得所有分類和問題
Dim rsCategories
Set rsCategories = conn.Execute("SELECT * FROM QuestionCategories ORDER BY SortOrder")

If Err.Number <> 0 Then
    Response.Write "資料庫錯誤: " & Err.Description
    Response.End
End If
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
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <main class="main-content">
            <header class="top-bar">
                <div class="search-bar">
                    <input type="search" placeholder="輸入公司名稱模糊搜尋...">
                </div>
                <div class="user-actions">
                    <!-- 移除這裡的 theme_switch.asp include -->
                </div>
            </header>

            <div class="visit-form-container">
                <h1>訪廠記錄表</h1>
                
                <form id="visitForm" action="save_visit.asp" method="post">
                    <div class="company-info">
                        <div class="form-group">
                            <label for="companyName">公司名稱</label>
                            <select id="companyName" name="companyName" required>
                                <option value="">請選擇公司</option>
                                <% 
                                Dim rsVendors
                                Set rsVendors = conn.Execute("SELECT VendorID, VendorName FROM Vendors WHERE IsActive = 1 ORDER BY VendorName")
                                Do While Not rsVendors.EOF 
                                %>
                                    <option value="<%=rsVendors("VendorName")%>"><%=rsVendors("VendorName")%></option>
                                <%
                                    rsVendors.MoveNext
                                Loop
                                %>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="visitDate">訪談日期</label>
                            <input type="date" id="visitDate" name="visitDate" required
                                   value="<%=Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)%>">
                        </div>
                    </div>

                    <div class="questions-container">
                        <%
                        If Not rsCategories.EOF Then
                            Do While Not rsCategories.EOF
                                Dim categoryID
                                categoryID = rsCategories("CategoryID")
                        %>
                            <div class="category-section">
                                <h2><%=rsCategories("CategoryName")%>
                                <% If Not rsCategories("IsRequired") Then %>
                                    <span class="optional-tag">選填</span>
                                <% End If %>
                                </h2>

                                <div class="questions">
                                    <% 
                                    Dim rsQuestions
                                    Set rsQuestions = conn.Execute("SELECT * FROM VisitQuestions WHERE CategoryID = " & categoryID & " ORDER BY SortOrder")
                                    
                                    Do While Not rsQuestions.EOF 
                                    %>
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
                                                    Dim optionItems, optionItem
                                                    optionItems = Split(Replace(Replace(rsQuestions("Options"), "[", ""), "]", ""), ",")
                                                    For Each optionItem in optionItems
                                                    %>
                                                        <option value="<%=Replace(Replace(optionItem,"""","")," ","")%>">
                                                            <%=Replace(Replace(optionItem,"""","")," ","")%>
                                                        </option>
                                                    <% Next %>
                                                </select>   
                                            <% Else %>
                                                <input type="text" name="q_<%=rsQuestions("QuestionID")%>"
                                                       <%=IIf(rsQuestions("IsRequired"), "required", "")%>>
                                            <% End If %>

                                            <% If rsQuestions("CanModify") Then %>
                                                <button type="button" onclick="editQuestion(<%=rsQuestions("QuestionID")%>)">
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
                        End If
                        %>
                    </div>
                </form>
            </div>
        </main>
    </div>

    <script>
        function editQuestion(questionId) {
            const inputElement = document.querySelector(`[name="q_${questionId}"]`);
            const companySelect = document.getElementById('companyName');
            const visitDateInput = document.getElementById('visitDate');
            
            if (!inputElement || !companySelect || !visitDateInput) {
                console.error('找不到必要的表單元素');
                return;
            }

            const answer = inputElement.value;
            const companyName = companySelect.value;
            const visitDate = visitDateInput.value;
            
            if (!companyName) {
                alert('請先選擇公司名稱');
                return;
            }
            
            if (!answer) {
                alert('請輸入答案');
                return;
            }

            const formData = new URLSearchParams();
            formData.append('questionId', questionId);
            formData.append('companyName', companyName);
            formData.append('answer', answer);
            formData.append('visitDate', visitDate);

            fetch('save_answer.asp', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: formData.toString()
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    sessionStorage.setItem('selectedCompany', companyName);
                    sessionStorage.setItem('visitDate', visitDate);
                    window.location.reload();
                } else {
                    alert(data.message || '儲存失敗');
                }
            })
            .catch(error => {
                console.error('Error:', error);
                alert('儲存時發生錯誤');
            });
        }

        window.addEventListener('load', function() {
            const selectedCompany = sessionStorage.getItem('selectedCompany');
            const visitDate = sessionStorage.getItem('visitDate');
            
            if (selectedCompany) {
                const companySelect = document.getElementById('companyName');
                companySelect.value = selectedCompany;
                companySelect.dispatchEvent(new Event('change'));
            }
            
            if (visitDate) {
                const visitDateInput = document.getElementById('visitDate');
                visitDateInput.value = visitDate;
                sessionStorage.removeItem('visitDate');
            }
            
            sessionStorage.removeItem('selectedCompany');
        });

        document.getElementById('companyName').addEventListener('change', function() {
            const companyName = this.value;
            if (!companyName) return;

            // 清除所有現有的最近答案顯示
            document.querySelectorAll('.last-answer').forEach(el => el.remove());

            // 獲取所有問題的最近答案
            fetch(`get_last_answers.asp?companyName=${encodeURIComponent(companyName)}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        // 為每個問題添加最近答案
                        data.answers.forEach(answer => {
                            const questionInput = document.querySelector(`[name="q_${answer.QuestionID}"]`);
                            if (questionInput) {
                                // 創建最近答案的顯示元素
                                const lastAnswerDiv = document.createElement('div');
                                lastAnswerDiv.className = 'last-answer';
                                
                                // 創建日期 badge
                                const dateBadge = document.createElement('span');
                                dateBadge.className = 'date-badge';
                                dateBadge.textContent = answer.ModifiedDate;
                                
                                // 創建答案文字元素
                                const answerText = document.createElement('div');
                                answerText.textContent = '回答：' + answer.Answer;
                                
                                // 組合元素
                                lastAnswerDiv.appendChild(dateBadge);
                                lastAnswerDiv.appendChild(answerText);
                                
                                // 插入到問題輸入框之前
                                questionInput.parentNode.insertBefore(lastAnswerDiv, questionInput);
                            }
                        });
                    }
                })
                .catch(error => {
                    console.error('Error fetching last answers:', error);
                });
        });

        // 在 script 區塊中添加公司名稱搜尋功能
        document.querySelector('.search-bar input').addEventListener('input', function(e) {
            const searchText = e.target.value.toLowerCase().trim();
            const companySelect = document.getElementById('companyName');
            const options = companySelect.options;
            
            // 從第二個選項開始遍歷（跳過"請選擇"選項）
            for (let i = 1; i < options.length; i++) {
                const optionText = options[i].text.toLowerCase();
                // 如果搜尋文字為空或選項文字包含搜尋文字，則顯示該選項
                if (searchText === '' || optionText.includes(searchText)) {
                    options[i].style.display = '';
                } else {
                    options[i].style.display = 'none';
                }
            }
        });

        // 修改 select 的樣式，使隱藏的選項在下拉時真的隱藏
        document.getElementById('companyName').addEventListener('mousedown', function(e) {
            if (e.target.tagName === 'OPTION' && e.target.style.display === 'none') {
                e.preventDefault();
            }
        });
    </script>
</body>
</html> 