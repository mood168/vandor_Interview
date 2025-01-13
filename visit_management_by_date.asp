<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 定義ADODB常數
Const adDate = 7
Const adParamInput = 1
Const adVarChar = 200

' 日期格式化函數
Function FormatDateYMD(dateValue)
    If IsNull(dateValue) Or dateValue = "" Then
        FormatDateYMD = "-"
    Else
        FormatDateYMD = Year(dateValue) & "/" & Right("0" & Month(dateValue), 2) & "/" & Right("0" & Day(dateValue), 2)
    End If
End Function

' 取得訪廠記錄列表
Dim sql
sql = "WITH RankedVisits AS ( " & _
      "    SELECT " & _
      "        vr.VisitorID, vr.VisitID, " & _
      "        vr.CompanyName, " & _
      "        ISNULL(v.ParentCode, '') as ParentCode, " & _
      "        ISNULL(v.ChildCode, '') as ChildCode, " & _
      "        ISNULL(vr.Interviewee, '') as Interviewee, " & _
      "        vr.VisitDate, " & _
      "        vr.Status, " & _
      "        ISNULL(u.FullName, '') as VisitorName, " & _
      "        ISNULL((SELECT TOP 1 ModifiedDate FROM VisitAnswers va " & _
      "         WHERE va.VisitID = vr.VisitID ORDER BY ModifiedDate DESC), vr.CreatedDate) as LastAnswerDate, " & _
      "        ROW_NUMBER() OVER (PARTITION BY vr.CompanyName ORDER BY vr.VisitDate DESC) as RowNum, " & _
      "        COUNT(*) OVER() as TotalCount " & _
      "    FROM VisitRecords vr " & _
      "    LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
      "    LEFT JOIN Vendors v ON vr.CompanyName = v.VendorName " & _
      "    WHERE 1=1 "

' 取得日期參數
Dim startDate, endDate
startDate = Request.QueryString("startDate")
endDate = Request.QueryString("endDate")

' 根據日期參數動態添加條件
If startDate <> "" Then
    sql = sql & " AND vr.VisitDate >= '" & startDate & "'"
End If

If endDate <> "" Then
    sql = sql & " AND vr.VisitDate <= '" & endDate & "'"
End If

sql = sql & " ) SELECT * FROM RankedVisits WHERE RowNum = 1 ORDER BY VisitDate DESC"

' 執行查詢
Dim rs
Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>訪廠紀錄查詢</title>
    <link rel="stylesheet" href="styles/dashboard.css">
    <link rel="stylesheet" href="styles/visit_management.css">
    <style>
        .search-bar select {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            margin-right: 10px;
        }
        
        .result-count {
            margin: 10px 0;
            color: #666;
            font-size: 14px;
        }
        
        #visitorFilter {
            min-width: 150px;
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
                    <input type="date" 
                           id="startDate" 
                           placeholder="開始日期"
                           title="選擇開始日期" class="save-btn">
                    <span>~</span>
                    <input type="date"
                           id="endDate"
                           placeholder="結束日期" 
                           title="選擇結束日期" class="save-btn">
                    
                    <select id="visitorFilter" title="選擇訪廠人員" class="save-btn">
                        <option value="">全部人員</option>
                        <% 
                        Dim rsVisitors
                        Set rsVisitors = conn.Execute("SELECT DISTINCT u.UserID, u.FullName FROM Users u INNER JOIN VisitRecords vr ON u.UserID = vr.VisitorID WHERE u.IsActive = 1 ORDER BY u.FullName")
                        Do While Not rsVisitors.EOF 
                        %>
                            <option value="<%=rsVisitors("UserID")%>"><%=rsVisitors("FullName")%></option>
                        <%
                            rsVisitors.MoveNext
                        Loop
                        %>
                    </select>

                    <input type="text" 
                           id="parentCodeFilter" 
                           placeholder="母代號"
                           title="輸入母代號搜尋" 
                           class="save-btn" style="width: 100px;">
                    
                    <input type="text" 
                           id="childCodeFilter" 
                           placeholder="子代號"
                           title="輸入子代號搜尋" 
                           class="save-btn" style="width: 100px;">
                    
                    <button type="button" id="searchBtn" class="save-btn">搜尋</button>
                </div>

                <div class="result-count">
                    查詢結果: <span id="resultCount">0</span> 筆
                </div>
            </header>

            <div class="content">
                <h1>訪廠紀錄查詢</h1>
                
                <div class="visits-table-container">
                    <table class="visits-table">
                        <thead>
                            <tr>
                                <th>公司名稱</th>
                                <th>母代號</th>
                                <th>子代號</th>
                                <th>訪廠人員</th>
                                <th>受訪人</th>
                                <th>訪廠日期</th>
                                <th>最後更新</th>
                                <th>狀態</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rs.EOF %>
                                <tr data-visitor-id="<%=rs("VisitorID")%>">
                                    <td><%=rs("CompanyName")%></td>
                                    <td><%=rs("ParentCode")%></td>
                                    <td><%=rs("ChildCode")%></td>
                                    <td><%=rs("VisitorName")%></td>
                                    <td><%=rs("Interviewee")%></td>
                                    <td><%=FormatDateYMD(rs("VisitDate"))%></td>
                                    <td><%=FormatDateYMD(rs("LastAnswerDate"))%></td>
                                    <td>
                                        <span class="status-badge <%=LCase(rs("Status"))%>">
                                            <%=rs("Status")%>
                                        </span>
                                    </td>
                                    <td class="actions">
                                        <a href="visit_questions.asp?vendor=<%=rs("CompanyName")%>" 
                                           class="edit-btn">編輯</a>
                                        <a href="print_visit.asp?id=<%=rs("VisitID")%>" 
                                           class="edit-btn" target="_blank">訪廠紀錄表</a>
                                        <a href="full_visit_records_by_date_range.asp?id=<%=rs("VisitorID")%>" 
                                           class="edit-btn" target="_blank">完整紀錄表</a>
                                    </td>
                                </tr>
                            <% 
                                rs.MoveNext
                                Loop 
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </main>
    </div>

    <script>
        // 初始化結果數量
        window.addEventListener('load', function() {
            const visibleRows = document.querySelectorAll('.visits-table tbody tr:not([style*="display: none"])').length;
            document.getElementById('resultCount').textContent = visibleRows;
        });

        // 日期範圍搜尋功能
        document.getElementById('searchBtn').addEventListener('click', function() {
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const selectedVisitor = document.getElementById('visitorFilter').value;
            const parentCode = document.getElementById('parentCodeFilter').value.toLowerCase().trim();
            const childCode = document.getElementById('childCodeFilter').value.toLowerCase().trim();
            
            // 驗證日期邏輯
            if (startDate && endDate && startDate > endDate) {
                alert('開始日期不能大於結束日期');
                return;
            }
            
            const rows = document.querySelectorAll('.visits-table tbody tr');
            let visibleCount = 0;
            
            rows.forEach(row => {
                const visitDate = row.cells[5].textContent; // 訪廠日期索引從3改為5
                const visitorId = row.getAttribute('data-visitor-id'); // 訪廠人員ID
                const rowParentCode = row.cells[1].textContent.toLowerCase(); // 母代號
                const rowChildCode = row.cells[2].textContent.toLowerCase(); // 子代號
                
                // 日期篩選邏輯
                let dateMatch = true;
                if (startDate && endDate) {
                    const visit = new Date(visitDate);
                    const start = new Date(startDate);
                    const end = new Date(endDate);
                    
                    // 設定時間為00:00:00以確保正確比較
                    start.setHours(0,0,0,0);
                    end.setHours(23,59,59,999);
                    visit.setHours(12,0,0,0); // 設定訪廠日期為中午以避免時區問題
                    
                    dateMatch = visit >= start && visit <= end;
                }
                
                // 訪廠人員篩選邏輯
                const visitorMatch = !selectedVisitor || visitorId === selectedVisitor;
                
                // 代號篩選邏輯
                const parentCodeMatch = !parentCode || rowParentCode.includes(parentCode);
                const childCodeMatch = !childCode || rowChildCode.includes(childCode);
                
                // 組合所有篩選條件
                const shouldShow = dateMatch && visitorMatch && parentCodeMatch && childCodeMatch;
                row.style.display = shouldShow ? '' : 'none';
                
                if (shouldShow) {
                    visibleCount++;
                }
            });
            
            // 更新結果數量顯示
            document.getElementById('resultCount').textContent = visibleCount;
        });
    </script>

    <!-- 在表格之後，body 結束前添加 Modal -->
    <div id="editVisitModal" class="modal">
        <div class="modal-content">
            <h2>編輯訪廠記錄</h2>
            <form id="editVisitForm" onsubmit="return saveVisit(event)">
                <input type="hidden" id="editVisitId" name="editVisitId">
                <div class="form-group">
                    <label for="editCompanyName">公司名稱</label>
                    <input type="text" id="editCompanyName" name="editCompanyName" required>
                </div>
                <div class="form-group">
                    <label for="editVisitorId">訪廠人員</label>
                    <select id="editVisitorId" name="editVisitorId" required>
                        <% 
                        Dim rsUsers
                        Set rsUsers = conn.Execute("SELECT UserID, FullName FROM Users WHERE IsActive = 1")
                        Do While Not rsUsers.EOF 
                        %>
                            <option value="<%=rsUsers("UserID")%>"><%=rsUsers("FullName")%></option>
                        <%
                            rsUsers.MoveNext
                        Loop
                        %>
                    </select>
                </div>
                <div class="form-group">
                    <label for="editInterviewee">受訪人</label>
                    <input type="text" id="editInterviewee" name="editInterviewee" value>
                </div>
                <div class="form-group">
                    <label for="editVisitDate">訪廠日期</label>
                    <input type="date" id="editVisitDate" name="editVisitDate" required>
                </div>
                <div class="form-group">
                    <label for="editStatus">狀態</label>
                    <select id="editStatus" name="editStatus" required>
                        <option value="Draft">草稿</option>
                        <option value="Completed">完成</option>
                        <option value="Reviewed">已審核</option>
                    </select>
                </div>
                <div class="modal-actions">
                    <button type="submit" class="save-btn">儲存</button>
                    <button type="button" class="cancel-btn" onclick="hideEditVisitModal()">取消</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    // 編輯訪廠記錄
    function editVisit(visitId) {
        fetch(`get_visit.asp?id=${visitId}`)
            .then(response => response.text())
            .then(text => {
                try {
                    console.log('Server response:', text); // 除錯用
                    const data = JSON.parse(text);
                    if (data.success) {
                        document.getElementById('editVisitId').value = data.VisitID;
                        document.getElementById('editCompanyName').value = data.CompanyName;
                        document.getElementById('editCompanyName').readOnly = true;
                        document.getElementById('editVisitorId').value = data.VisitorID;
                        document.getElementById('editInterviewee').value = data.Interviewee || '';
                        document.getElementById('editVisitDate').value = data.VisitDate;
                        document.getElementById('editStatus').value = data.Status;
                        
                        console.log('Loaded data:', { // 除錯用
                            visitId: data.VisitID,
                            companyName: data.CompanyName,
                            visitorId: data.VisitorID,
                            interviewee: data.Interviewee,
                            visitDate: data.VisitDate,
                            status: data.Status
                        });
                        
                        showEditVisitModal();
                    } else {
                        alert(data.message);
                    }
                } catch (error) {
                    console.error('JSON parse error:', error);
                    console.error('Raw response:', text);
                    alert('載入資料時發生錯誤');
                }
            })
            .catch(error => {
                console.error('Fetch error:', error);
                alert('載入訪廠記錄時發生錯誤');
            });
    }

    function showEditVisitModal() {
        document.getElementById('editVisitModal').style.display = 'flex';
    }

    function hideEditVisitModal() {
        document.getElementById('editVisitModal').style.display = 'none';
    }

    function saveVisit(event) {
        event.preventDefault();
        
        // 取得表單元素
        const form = document.getElementById('editVisitForm');
        
        // 直接使用表單建立 FormData
        const formData = new FormData(form);
        
        // 檢查必要欄位
        const visitId = formData.get('editVisitId');
        const companyName = formData.get('editCompanyName');
        const visitorId = formData.get('editVisitorId');
        const visitDate = formData.get('editVisitDate');
        const status = formData.get('editStatus');
        
        // 驗證必填欄位
        if (!visitId && !companyName) {
            alert('請輸入公司名稱');
            return false;
        }
        if (!visitorId) {
            alert('請選擇訪廠人員');
            return false;
        }
        if (!visitDate) {
            alert('請選擇訪廠日期');
            return false;
        }
        if (!status) {
            alert('請選擇狀態');
            return false;
        }
        
        // 將 FormData 轉換為 URL 編碼字串
        const urlEncodedData = new URLSearchParams(formData).toString();
        
        // 發送請求
        fetch('save_visit.asp', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: urlEncodedData
        })
        .then(response => response.text())
        .then(text => {
            try {
                console.log('Server response:', text); // 除錯用
                const data = JSON.parse(text);
                if (data.success) {
                    alert('訪廠記錄已更新');
                    window.location.href = 'visit_management.asp';  // 或其他目標頁面
                } else {
                    alert(data.message || '儲存失敗');
                }
            } catch (error) {
                console.error('回應內容:', text);
                alert('處理回應時發生錯誤：' + error.message);
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('儲存時發生錯誤：' + error.message);
        });
        
        return false;
    }
    </script>
</body>
</html> 