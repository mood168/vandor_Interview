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
    <style>
        /* 在 head 區塊中加入或修改以下樣式 */
        .top-bar {
            display: flex;
            justify-content: center; /* 水平置中 */
            align-items: center;    /* 垂直置中 */
            padding: 1rem;
            background-color: var(--header-bg);
        }

        .search-bar {
            width: 100%;
            max-width: 600px;      /* 限制最大寬度 */
            margin: 0 auto;        /* 水平置中 */
        }

        .search-bar input {
            width: 100%;
            padding: 10px 15px;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            font-size: 16px;
            background-color: var(--input-bg);
            color: var(--text-color);
        }
        
        .save-btn {
            padding: 8px 16px;
            border: 2px solid #4a90e2;  /* 藍色外框 */
            border-radius: 4px;
            background-color: white;
            color: #4a90e2;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-left: 10px;
        }

        .save-btn:hover {
            background-color: #4a90e2;
            color: white;
        }

        .save-btn:active {
            transform: translateY(1px);
        }

        .user-actions {
            display: none;         /* 隱藏不需要的使用者操作區 */
        }

        /* Radio 按鈕樣式 */
        .radio-label input[type="radio"] {
            margin-right: 5px;
        }

        .radio-label:hover {
            border: 1px solid #4a90e2;
            border-radius: 4px;
            padding: 4px 8px;
            background-color: #f5f8ff;
        }

        .radio-label input[type="radio"]:checked + span {
            color: #4a90e2;
            font-weight: bold;
        }

        .radio-label input[type="radio"]:focus + span {
            outline: 2px solid #4a90e2;
            outline-offset: 2px;
            border-radius: 2px;
        }

        /* Checkbox 按鈕樣式 */
        .checkbox-label input[type="checkbox"] {
            margin-right: 5px;
        }

        .checkbox-label:hover {
            border: 1px solid #4a90e2;
            border-radius: 4px;
            padding: 4px 8px;
            background-color: #f5f8ff;
        }

        .checkbox-label input[type="checkbox"]:checked + span {
            color: #4a90e2;
            font-weight: bold;
        }

        .checkbox-label input[type="checkbox"]:focus + span {
            outline: 2px solid #4a90e2;
            outline-offset: 2px;
            border-radius: 2px;
        }

        /* 確保 label 有足夠的間距 */
        .radio-label, .checkbox-label {
            margin: 10px;
            display: inline-block;
            position: relative;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .checkbox-group {
            display: flex;
            flex-wrap: wrap;  /* 允許選項換行 */
            gap: 10px;        /* 選項之間的間距 */
            width: 100%;
            max-height: none; /* 移除最大高度限制 */
            overflow: visible; /* 移除 overflow 設定 */
        }

        .checkbox-label {
            flex: 0 1 auto;   /* 允許元素根據內容自動調整大小 */
            white-space: nowrap; /* 防止文字換行 */
            margin: 5px;
            padding: 4px 8px;
            border: 1px solid transparent; /* 預設透明邊框 */
        }

        /* 確保輸入框和標籤在同一行 */
        .checkbox-label input[type="number"] {
            width: 80px;      /* 設定合適的寬度 */
            margin: 0 5px;
            display: inline-block;
            vertical-align: middle;
        }

        /* 百分比輸入框的特定樣式 */
        .percentage-input {
            width: 60px !important;
        }

        /* 金額輸入框的特定樣式 */
        .amount-input {
            width: 100px !important;
        }
    </style>
</head>
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <main class="main-content">
            <header class="top-bar">
                <div class="search-bar">
                    <input type="search" placeholder="輸入公司名稱模糊搜尋..." value="<%= Request.QueryString("vendor") %>">
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
                                <h2><%=rsCategories("CategoryName")%></h2>

                                <div class="questions">
                                    <% 
                                    Dim rsQuestions
                                    Set rsQuestions = conn.Execute("SELECT * FROM VisitQuestions WHERE CategoryID = " & categoryID & " ORDER BY SortOrder")
                                    
                                    Do While Not rsQuestions.EOF 
                                        Dim questionId, answerType, hasOptions, options, hasPercentage
                                        questionId = rsQuestions("QuestionID")
                                        answerType = rsQuestions("AnswerType")
                                        hasOptions = rsQuestions("HasOptions")
                                        options = rsQuestions("Options")
                                        hasPercentage = rsQuestions("HasPercentage")
                                    %>
                                        <div class="question-item">
                                            <label style="font-size: 18px;">
                                                <%=rsQuestions("QuestionText")%>
                                                <% If rsQuestions("IsRequired") Then %>
                                                    <span class="required">*</span>
                                                <% End If %>
                                            </label>
                                            <hr/>
                                            <% 
                                            Select Case answerType
                                                Case "text" 
                                            %>
                                                    <input type="text" name="q_<%=questionId%>" 
                                                        <%=IIf(rsQuestions("IsRequired"), "required", "")%>>
                                            <% 
                                                Case "number" 
                                            %>
                                                    <input type="number" name="q_<%=questionId%>" 
                                                        <%=IIf(rsQuestions("IsRequired"), "required", "")%>>
                                            <% 
                                                Case "date" 
                                            %>
                                                    <input type="date" name="q_<%=questionId%>" 
                                                        <%=IIf(rsQuestions("IsRequired"), "required", "")%>>
                                            <% 
                                                Case "radio" 
                                                    If hasOptions Then
                                                        Dim radioOptions
                                                        radioOptions = Split(Replace(Replace(options, "[", ""), "]", ""), ",")
                                            %>
                                                    <div class="radio-group">
                                                        <% 
                                                            For Each opt in radioOptions 
                                                                opt = Replace(Replace(opt, """", ""), " ", "")
                                                        %>
                                                            <label class="radio-label" style="display: inline-block;">
                                                                <input type="radio" name="q_<%=questionId%>" value="<%=opt%>" <%=IIf(rsQuestions("IsRequired"), "required", "")%>><span><%=opt%></span>
                                                                <% If hasPercentage And InStr(opt, "佔") > 0 Then %>
                                                                    <input type="number" class="percentage-input" name="q_<%=questionId%>_percent_<%=opt%>" min="0" max="100" placeholder="%">
                                                                <% End If %>
                                                            </label>
                                                        <% Next %>
                                                    </div>
                                            <% 
                                                    End If

                                                Case "checkbox"
                                                    If hasOptions Then
                                                        Dim checkOptions
                                                        checkOptions = Split(Replace(Replace(options, "[", ""), "]", ""), ",")
                                            %>
                                                    <div class="checkbox-group">
                                                        <% 
                                                            For Each opt in checkOptions 
                                                                opt = Replace(Replace(opt, """", ""), " ", "")
                                                        %>
                                                            <label class="checkbox-label" style="display: inline-block;">
                                                                <input type="checkbox" name="q_<%=questionId%>" 
                                                                    value="<%=opt%>">
                                                                <span><%=opt%></span>
                                                                <% 
                                                                If hasPercentage And InStr(opt, "佔") > 0 Then 
                                                                    Dim inputName
                                                                    If InStr(opt, ",") > 0 Then
                                                                        ' 處理有金額的選項
                                                                        Response.Write "<input type='number' class='amount-input' " & _
                                                                                    "name='q_" & questionId & "_amount_" & opt & "' " & _
                                                                                    "placeholder='元'>"
                                                                    End If
                                                                %>
                                                                    <input type="number" class="percentage-input" 
                                                                        name="q_<%=questionId%>_percent_<%=opt%>" 
                                                                        min="0" max="100" placeholder="%">
                                                                <% End If %>
                                                            </label>
                                                        <% Next %>
                                                    </div>
                                            <% 
                                                    End If

                                                Case Else
                                            %>
                                                    <input type="text" name="q_<%=questionId%>" 
                                                        <%=IIf(rsQuestions("IsRequired"), "required", "")%>>
                                            <%
                                            End Select 
                                            %>

                                            <% If rsQuestions("CanModify") Then %>
                                                <button class="save-btn" type="button" onclick="saveAnswer(<%=questionId%>)">
                                                    儲存
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

        // 在頁面載入時執行
        window.addEventListener('load', function() {
            const urlParams = new URLSearchParams(window.location.search);
            const vendorFromUrl = urlParams.get('vendor');
            
            // 如果 URL 中有 vendor 參數
            if (vendorFromUrl) {
                const companySelect = document.getElementById('companyName');
                const decodedVendor = decodeURIComponent(vendorFromUrl);
                
                // 選中對應的選項
                for (let i = 0; i < companySelect.options.length; i++) {
                    if (companySelect.options[i].value === decodedVendor) {
                        companySelect.selectedIndex = i;
                        // 觸發 change 事件以載入最近答案
                        companySelect.dispatchEvent(new Event('change'));
                        break;
                    }
                }

                // 將公司名稱填入搜尋框
                const searchInput = document.querySelector('.search-bar input');
                if (searchInput) {
                    searchInput.value = decodedVendor;
                }
            }

            // 原有的 sessionStorage 相關代碼
            const selectedCompany = sessionStorage.getItem('selectedCompany');
            const visitDate = sessionStorage.getItem('visitDate');
            
            if (selectedCompany && !vendorFromUrl) {  // 只在沒有 URL 參數時才使用 sessionStorage
                const companySelect = document.getElementById('companyName');
                companySelect.value = selectedCompany;
                companySelect.dispatchEvent(new Event('change'));
            }
            
            if (visitDate) {
                const visitDateInput = document.getElementById('visitDate');
                visitDateInput.value = visitDate;
                sessionStorage.removeItem('visitDate');
            }
            
            if (!vendorFromUrl) {  // 只在沒有 URL 參數時才清除 sessionStorage
                sessionStorage.removeItem('selectedCompany');
            }
        });

        document.getElementById('companyName').addEventListener('change', function() {
            const companyName = this.value;
            if (!companyName) return;

            // 清除所有現的最近答案示
            document.querySelectorAll('.last-answer').forEach(el => el.remove());

            // 獲取所有問題的最近答案
            fetch(`get_last_answers.asp?companyName=${encodeURIComponent(companyName)}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        data.answers.forEach(answer => {
                            const questionInput = document.querySelector(`[name="q_${answer.QuestionID}"]`);
                            if (questionInput) {
                                const lastAnswerDiv = document.createElement('div');
                                lastAnswerDiv.className = 'last-answer';
                                
                                const dateBadge = document.createElement('span');
                                dateBadge.className = 'date-badge';
                                dateBadge.textContent = answer.ModifiedDate;
                                
                                const answerText = document.createElement('div');
                                answerText.textContent = '回答：' + answer.Answer;
                                
                                lastAnswerDiv.appendChild(dateBadge);
                                lastAnswerDiv.appendChild(answerText);
                                
                                if (questionInput.type === 'radio' || questionInput.type === 'checkbox') {
                                    const inputGroup = questionInput.closest('.radio-group, .checkbox-group');
                                    if (inputGroup) {
                                        inputGroup.parentNode.insertBefore(lastAnswerDiv, inputGroup.nextSibling);
                                    }
                                } else {
                                    questionInput.parentNode.insertBefore(lastAnswerDiv, questionInput.nextSibling);
                                }
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

        // 修改 select 的樣式使隱藏的選項在下拉時真的隱藏
        document.getElementById('companyName').addEventListener('mousedown', function(e) {
            if (e.target.tagName === 'OPTION' && e.target.style.display === 'none') {
                e.preventDefault();
            }
        });

        function saveAnswer(questionId) {
            console.log('questionId:', questionId); // 加入除錯訊息

            const form = document.getElementById('visitForm');
            const companyName = form.companyName.value;
            const visitDate = form.visitDate.value;
            
            if (!companyName) {
                alert('請選擇公司名稱');
                return;
            }

            // 收集答案數據
            const formData = new URLSearchParams();
            formData.append('questionId', questionId);  // 確保這裡有值
            formData.append('companyName', companyName);
            formData.append('visitDate', visitDate);

            let answer = '';
            
            try {
                // 根據問題類型收集答案
                const questionContainer = document.querySelector(`[name="q_${questionId}"]`).closest('.question-item');
                const inputElement = questionContainer.querySelector(`[name="q_${questionId}"]`);
                const answerType = inputElement.type;

                if (answerType === 'radio') {
                    const selectedRadio = questionContainer.querySelector(`input[name="q_${questionId}"]:checked`);
                    if (selectedRadio) {
                        answer = selectedRadio.value;
                        // 檢查是否有百分比輸入
                        const percentInput = questionContainer.querySelector(`[name="q_${questionId}_percent_${selectedRadio.value}"]`);
                        if (percentInput && percentInput.value) {
                            answer += `${percentInput.value}%`;
                        }
                    }
                } else if (answerType === 'checkbox') {
                    const checkedBoxes = questionContainer.querySelectorAll(`input[name="q_${questionId}"]:checked`);
                    const answers = [];
                    checkedBoxes.forEach(box => {
                        let value = box.value;
                        // 檢查是否有百分比和金額輸入
                        const percentInput = questionContainer.querySelector(`[name="q_${questionId}_percent_${box.value}"]`);
                        const amountInput = questionContainer.querySelector(`[name="q_${questionId}_amount_${box.value}"]`);
                        if (percentInput && percentInput.value) {
                            value += `${percentInput.value.replace('%', '')}%`;
                        }
                        if (amountInput && amountInput.value) {
                            value += `${amountInput.value.replace('%', '')}元`;
                        }
                        answers.push(value);
                    });
                    answer = answers.join(',');
                } else {
                    answer = inputElement.value;
                }

                if (!answer) {
                    alert('請輸入答案');
                    return;
                }

                formData.append('answer', answer);

                // 發送請求
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
                        alert('儲存成功');
                        // 重新載入最近答案
                        const companySelect = document.getElementById('companyName');
                        companySelect.dispatchEvent(new Event('change'));
                    } else {
                        alert(data.message || '儲存失敗');
                    }
                })
                .catch(error => {
                    console.error('Error:', error);
                    alert('儲存時發生錯誤');
                });
            } catch (error) {
                console.error('Error processing answer:', error);
                alert('處理答案時發生錯誤');
            }
        }
    </script>
</body>
</html> 