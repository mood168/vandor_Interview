<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 檢查是否為管理員
If Session("UserRole") <> "Admin" Then
    Response.Redirect "dashboard.asp"
    Response.End
End If

' 取得使用者列表
Dim rsUsers
Set rsUsers = conn.Execute("SELECT * FROM Users ORDER BY CreatedDate DESC")
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>使用者管理</title>
    <link rel="stylesheet" href="styles/dashboard.css">
    <link rel="stylesheet" href="styles/user_management.css">
</head>
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <main class="main-content">
            <header class="top-bar">
                <div class="search-bar">
                    <input type="search" id="userSearch" placeholder="搜尋使用者...">
                </div>
                <div class="user-actions">
                    <!--#include file="theme_switch.asp"-->
                    <button class="add-user-btn" onclick="showAddUserModal()">新增使用者</button>
                </div>
            </header>

            <div class="content">
                <h1>使用者管理</h1>
                
                <div class="users-table-container">
                    <table class="users-table">
                        <thead>
                            <tr>
                                <th>使用者名稱</th>
                                <th>姓名</th>
                                <th>部門</th>
                                <th>電話</th>
                                <th>Email</th>
                                <th>角色</th>
                                <th>狀態</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rsUsers.EOF %>
                                <tr>
                                    <td><%=rsUsers("Username")%></td>
                                    <td><%=rsUsers("FullName")%></td>
                                    <td><%=rsUsers("Department")%></td>
                                    <td><%=rsUsers("Phone")%></td>
                                    <td><%=rsUsers("Email")%></td>
                                    <td><%=rsUsers("UserRole")%></td>
                                    <td>
                                        <span class="status-badge <%
                                        If rsUsers("IsActive") Then
                                            Response.Write("active")
                                        Else
                                            Response.Write("inactive")
                                        End If
                                        %>">
                                            <%
                                        If rsUsers("IsActive") Then
                                            Response.Write("啟用")
                                        Else
                                            Response.Write("停用")
                                        End If
                                        %>
                                        </span>
                                    </td>
                                    <td class="actions">
                                        <button onclick="editUser(<%=rsUsers("UserID")%>)" class="edit-btn">編輯</button>                                        
                                    </td>
                                </tr>
                            <% 
                                rsUsers.MoveNext
                                Loop 
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </main>
    </div>

    <!-- 新增使用者 Modal -->
    <div id="addUserModal" class="modal">
        <div class="modal-content">
            <h2>新增使用者</h2>
            <form id="addUserForm" action="save_user.asp" method="post">
                <div class="form-group">
                    <label for="newUsername">使用者名稱</label>
                    <input type="text" id="newUsername" name="username" required>
                </div>
                <div class="form-group">
                    <label for="newPassword">密碼</label>
                    <input type="password" id="newPassword" name="password" required>
                </div>
                <div class="form-group">
                    <label for="newFullName">姓名</label>
                    <input type="text" id="newFullName" name="fullName" required>
                </div>
                <div class="form-group">
                    <label for="newPhone">電話</label>
                    <input type="tel" id="newPhone" name="phone">
                </div>
                <div class="form-group">
                    <label for="newEmail">Email</label>
                    <input type="email" id="newEmail" name="email">
                </div>
                <div class="form-group">
                    <label for="newDepartment">部門</label>
                    <input type="text" id="newDepartment" name="department">
                </div>
                <div class="form-group">
                    <label for="newUserRole">角色</label>
                    <select id="newUserRole" class="user-role-select" name="userRole" required>
                        <option value="User" class="user-role-option">一般使用者</option>
                        <option value="Manager" class="user-role-option">管理者</option>
                        <option value="Admin" class="user-role-option">系統管理員</option>
                    </select>
                </div>
                <div class="modal-actions">
                    <button type="submit" class="save-btn">儲存</button>
                    <button type="button" class="cancel-btn" onclick="hideAddUserModal()">取消</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 編輯使用者 Modal -->
    <div id="editUserModal" class="modal">
        <div class="modal-content">
            <h2>編輯使用者</h2>
            <form id="editUserForm" action="save_user.asp" method="post">
                <input type="hidden" id="editUserId" name="userId">
                <div class="form-group">
                    <label for="editUsername">使用者名稱</label>
                    <input type="text" id="editUsername" name="username" required>
                </div>
                <div class="form-group">
                    <label for="editPassword">密碼 (若不修改請留空)</label>
                    <input type="password" id="editPassword" name="password">
                </div>
                <div class="form-group">
                    <label for="editFullName">姓名</label>
                    <input type="text" id="editFullName" name="fullName" required>
                </div>
                <div class="form-group">
                    <label for="editPhone">電話</label>
                    <input type="tel" id="editPhone" name="phone">
                </div>
                <div class="form-group">
                    <label for="editEmail">Email</label>
                    <input type="email" id="editEmail" name="email">
                </div>
                <div class="form-group">
                    <label for="editDepartment">部門</label>
                    <input type="text" id="editDepartment" name="department">
                </div>
                <div class="form-group">
                    <label for="editUserRole">角色</label>
                    <select id="editUserRole" class="user-role-select" name="userRole" required>
                        <option value="User" class="user-role-option">一般使用者</option>
                        <option value="Manager" class="user-role-option">管理者</option>
                        <option value="Admin" class="user-role-option">系統管理員</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="editUserStatus">狀態</label>
                    <select id="editUserStatus" class="user-status-select" name="userStatus" required>
                        <option value="1" class="user-status-option">啟用</option>
                        <option value="0" class="user-status-option">停用</option>
                    </select>
                </div>
                <div class="modal-actions">
                    <button type="submit" class="save-btn">儲存</button>
                    <button type="button" class="cancel-btn" onclick="hideEditUserModal()">取消</button>
                </div>
            </form>
        </div>
    </div>

    <script>
        // Modal 控制
        function showAddUserModal() {
            document.getElementById('addUserModal').style.display = 'flex';
        }

        function hideAddUserModal() {
            document.getElementById('addUserModal').style.display = 'none';
        }

        // 使用者搜尋
        document.getElementById('userSearch').addEventListener('input', function(e) {
            const searchText = e.target.value.toLowerCase();
            const rows = document.querySelectorAll('.users-table tbody tr');
            
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(searchText) ? '' : 'none';
            });
        });

        // 編輯使用者
        function editUser(userId) {
            fetch(`get_user.asp?id=${userId}`)
                .then(response => {
                    if (!response.ok) {
                        throw new Error('Network response was not ok');
                    }
                    return response.json();
                })
                .then(data => {
                    if (data.success) {
                        document.getElementById('editUserId').value = data.UserID;
                        document.getElementById('editUsername').value = data.Username;
                        document.getElementById('editFullName').value = data.FullName;
                        document.getElementById('editPhone').value = data.Phone || '';
                        document.getElementById('editEmail').value = data.Email || '';
                        document.getElementById('editDepartment').value = data.Department || '';
                        document.getElementById('editUserRole').value = data.UserRole;
                        document.getElementById('editUserStatus').value = data.IsActive ? "1" : "0";
                        
                        showEditUserModal();
                    } else {
                        alert(data.message);
                    }
                })
                .catch(error => {
                    console.error('Error fetching user data:', error);
                    alert('載入使用者資料時發生錯誤');
                });
        }

        function showEditUserModal() {
            document.getElementById('editUserModal').style.display = 'flex';
        }

        function hideEditUserModal() {
            document.getElementById('editUserModal').style.display = 'none';
        }
    </script>
</body>
</html> 