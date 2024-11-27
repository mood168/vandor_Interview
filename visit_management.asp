<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得訪廠記錄列表
Dim sql
sql = "SELECT " & _
      "vr.VisitID, " & _
      "vr.CompanyName, " & _
      "vr.Interviewee, " & _
      "vr.VisitDate, " & _
      "vr.Status, " & _
      "u.FullName as VisitorName, " & _
      "(SELECT TOP 1 ModifiedDate FROM VisitAnswers va " & _
      "WHERE va.VisitID = vr.VisitID ORDER BY ModifiedDate DESC) as LastAnswerDate " & _
      "FROM VisitRecords vr " & _
      "LEFT JOIN Users u ON vr.VisitorID = u.UserID " & _
      "ORDER BY vr.VisitDate DESC"

Dim rs
Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>訪廠管理</title>
    <link rel="stylesheet" href="styles/dashboard.css">
    <link rel="stylesheet" href="styles/visit_management.css">
</head>
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <main class="main-content">
            <header class="top-bar">
                <div class="search-bar">
                    <input type="search" 
                           id="visitSearch" 
                           placeholder="搜尋公司名稱、訪廠人員或受訪人..."
                           title="可輸入公司名稱、訪廠人員或受訪人進行搜尋">
                </div>
            </header>

            <div class="content">
                <h1>訪廠管理</h1>
                
                <div class="visits-table-container">
                    <table class="visits-table">
                        <thead>
                            <tr>
                                <th>公司名稱</th>
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
                                <tr>
                                    <td><%=rs("CompanyName")%></td>
                                    <td><%=rs("VisitorName")%></td>
                                    <td><%=rs("Interviewee")%></td>
                                    <td><%=FormatDateTime(rs("VisitDate"),2)%></td>
                                    <td><%
                                    If IsNull(rs("LastAnswerDate")) Then
                                        Response.Write("-")
                                    Else
                                        Response.Write(FormatDateTime(rs("LastAnswerDate"),2))
                                    End If
                                    %></td>
                                    <td>
                                        <span class="status-badge <%=LCase(rs("Status"))%>">
                                            <%=rs("Status")%>
                                        </span>
                                    </td>
                                    <td class="actions">
                                        <a href="#" 
                                           class="edit-btn" onclick="editVisit(<%=rs("VisitID")%>)">編輯</a>
                                        <a href="#" 
                                           class="edit-btn">訪廠紀錄表</a>
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
        // 搜尋功能
        document.getElementById('visitSearch').addEventListener('input', function(e) {
            const searchText = e.target.value.toLowerCase().trim();
            const rows = document.querySelectorAll('.visits-table tbody tr');
            
            rows.forEach(row => {
                const companyName = row.cells[0].textContent.toLowerCase(); // 公司名稱
                const visitorName = row.cells[1].textContent.toLowerCase(); // 訪廠人員
                const interviewee = row.cells[2].textContent.toLowerCase(); // 受訪人
                
                // 檢查是否符合任一搜尋條件
                const matchCompany = companyName.includes(searchText);
                const matchVisitor = visitorName.includes(searchText);
                const matchInterviewee = interviewee.includes(searchText);
                
                // 如果符合任一條件就顯示該列
                row.style.display = (matchCompany || matchVisitor || matchInterviewee) ? '' : 'none';
            });
        });
    </script>

    <!-- 在表格之後，body 結束前添加 Modal -->
    <div id="editVisitModal" class="modal">
        <div class="modal-content">
            <h2>編輯訪廠記錄</h2>
            <form id="editVisitForm" onsubmit="return saveVisit(event)">
                <input type="hidden" id="editVisitId" name="visitId">
                <div class="form-group">
                    <label for="editCompanyName">公司名稱</label>
                    <input type="text" id="editCompanyName" name="companyName" required>
                </div>
                <div class="form-group">
                    <label for="editVisitorId">訪廠人員</label>
                    <select id="editVisitorId" name="visitorId" required>
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
                    <input type="text" id="editInterviewee" name="interviewee">
                </div>
                <div class="form-group">
                    <label for="editVisitDate">訪廠日期</label>
                    <input type="date" id="editVisitDate" name="visitDate" required>
                </div>
                <div class="form-group">
                    <label for="editStatus">狀態</label>
                    <select id="editStatus" name="status" required>
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
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    document.getElementById('editVisitId').value = data.VisitID;
                    document.getElementById('editCompanyName').value = data.CompanyName;
                    document.getElementById('editVisitorId').value = data.VisitorID;
                    document.getElementById('editInterviewee').value = data.Interviewee || '';
                    document.getElementById('editVisitDate').value = data.VisitDate;
                    document.getElementById('editStatus').value = data.Status;
                    
                    showEditVisitModal();
                } else {
                    alert(data.message);
                }
            })
            .catch(error => {
                console.error('Error:', error);
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
        
        const formData = new FormData(document.getElementById('editVisitForm'));
        
        fetch('save_visit.asp', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                alert('訪廠記錄已更新');
                location.reload();
            } else {
                alert(data.message || '儲存失敗');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('儲存時發生錯誤');
        });
        
        return false;
    }
    </script>
</body>
</html> 