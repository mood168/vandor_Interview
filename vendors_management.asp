<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得廠商列表
Dim rsVendors
Set rsVendors = conn.Execute("SELECT * FROM Vendors ORDER BY IsActive DESC, CreatedDate DESC")
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>廠商管理</title>
    <link rel="stylesheet" href="styles/dashboard.css">
    <link rel="stylesheet" href="styles/vendors_management.css">
</head>
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <main class="main-content">
            <header class="top-bar">
                <div class="search-bar">
                    <input type="search" id="vendorSearch" 
                           placeholder="搜尋廠商 (可使用母代號、子代號、統一編號或廠商名稱搜尋)..."
                           title="可使用母代號、子代號、統一編號或廠商名稱搜尋" class="save-btn" style="width: 600px;">
                </div>
                <div class="user-actions">
                    <button class="add-vendor-btn" onclick="showAddVendorModal()">新增廠商</button>
                </div>
            </header>

            <div class="content">
                <h1>廠商管理</h1>
                
                <div class="vendors-table-container">
                    <table class="vendors-table">
                        <thead>
                            <tr>
                                <th>母代號</th>
                                <th>子代號</th>
                                <th>統一編號</th>
                                <th>廠商名稱</th>
                                <th>聯絡人</th>
                                <th>電話</th>
                                <th>地址</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rsVendors.EOF %>
                                <tr class="<% If Not rsVendors("IsActive") Then Response.Write "inactive" %>">
                                    <td><%=rsVendors("ParentCode")%></td>
                                    <td><%=rsVendors("ChildCode")%></td>
                                    <td><%=rsVendors("UniformNumber")%></td>
                                    <td><%=rsVendors("VendorName")%></td>
                                    <td><%=rsVendors("ContactPerson")%></td>
                                    <td><%=rsVendors("Phone")%></td>
                                    <td><%=rsVendors("Address")%></td>
                                    <td class="actions">
                                        
                                        
                                            <%
                                            If(rsVendors("IsActive") = True) then
                                            %>
                                            <button class="edit-btn" onclick="editVendor(<%=rsVendors("VendorID")%>)">編輯</button>
                                            <button class="delete-btn" onclick="deleteVendor(<%=rsVendors("VendorID")%>)">停用</button>
                                            <%
                                            Else
                                            %>
                                            <button class="activate-btn" onclick="activateVendor(<%=rsVendors("VendorID")%>)">啟用</button>
                                            <%
                                            End If
                                            %>
                                        
                                    </td>
                                </tr>
                            <% 
                                rsVendors.MoveNext
                                Loop 
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </main>
    </div>

    <!-- 新增廠商 Modal -->
    <div id="vendorModal" class="modal">
        <div class="modal-content">
            <h2>新增廠商</h2>
            <form id="vendorForm" action="save_vendor.asp" method="post">
                <input type="hidden" id="vendorId" name="vendorId">
                <div class="form-row">
                    <div class="form-group">
                        <label for="parentCode">母代號</label>
                        <input type="text" id="parentCode" name="parentCode" maxlength="3" required 
                               pattern="[A-Za-z0-9]{3}" title="請輸入3碼英數字">
                    </div>
                    <div class="form-group">
                        <label for="childCode">子代號</label>
                        <input type="text" id="childCode" name="childCode" maxlength="3" required
                               pattern="[A-Za-z0-9]{3}" title="請輸入3碼數字">
                    </div>
                </div>
                <div class="form-group">
                    <label for="uniformNumber">統一編號</label>
                    <input type="text" id="uniformNumber" name="uniformNumber" maxlength="8" required
                           pattern="[0-9]{8}" title="請輸入8碼數字">
                </div>
                <div class="form-group">
                    <label for="vendorName">廠商名稱</label>
                    <input type="text" id="vendorName" name="vendorName" maxlength="100" required>
                </div>
                <div class="form-group">
                    <label for="contactPerson">聯絡窗口</label>
                    <input type="text" id="contactPerson" name="contactPerson" maxlength="100" required>
                </div>
                <!--
                <div class="form-group">
                    <label for="logisticsContact">物流聯絡人</label>
                    <input type="text" id="logisticsContact" name="logisticsContact" maxlength="100">
                </div>
                <div class="form-group">
                    <label for="marketingContact">行銷聯絡人</label>
                    <input type="text" id="marketingContact" name="marketingContact" maxlength="100">
                </div>
                <div class="form-group">
                    <label for="phone">電話</label>
                    <input type="tel" id="phone" name="phone" maxlength="15">
                </div>
                -->
                <div class="form-group">
                    <label for="address">地址</label>
                    <input type="text" id="address" name="address" maxlength="100">
                </div>
                 <!--
                 <div class="form-group">
                    <label for="email">電子郵件</label>
                    <input type="email" id="email" name="email" maxlength="100">
                </div>
                -->
                <div class="form-group">
                    <label for="website">網址</label>
                    <input type="url" id="website" name="website" maxlength="250">
                </div>
                <div class="modal-actions">
                    <button type="submit" class="save-btn">儲存</button>
                    <button type="button" class="cancel-btn" onclick="hideVendorModal()">取消</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 編輯廠商 Modal -->
    <div id="editVendorModal" class="modal">
        <div class="modal-content">
            <h2>編輯廠商</h2>
            <form id="editVendorForm" action="save_vendor.asp" method="post">
                <input type="hidden" id="editVendorId" name="vendorId">
                <div class="form-row">
                    <div class="form-group">
                        <label for="editParentCode">母代號</label>
                        <input type="text" id="editParentCode" name="parentCode" maxlength="3" required 
                               pattern="[0-9]{3}" title="請輸入3碼數字">
                    </div>
                    <div class="form-group">
                        <label for="editChildCode">子代號</label>
                        <input type="text" id="editChildCode" name="childCode" maxlength="3" required
                               pattern="[0-9]{3}" title="請輸入3碼數字">
                    </div>
                </div>
                <div class="form-group">
                    <label for="editUniformNumber">統一編號</label>
                    <input type="text" id="editUniformNumber" name="uniformNumber" maxlength="8" required
                           pattern="[0-9]{8}" title="請輸入8碼數字">
                </div>
                <div class="form-group">
                    <label for="editVendorName">廠商名稱</label>
                    <input type="text" id="editVendorName" name="vendorName" maxlength="100" required>
                </div>
                <div class="form-group">
                    <label for="editContactPerson">聯絡窗口</label>
                    <input type="text" id="editContactPerson" name="contactPerson" maxlength="100" required>
                </div>
                <!--
                <div class="form-group">
                    <label for="editLogisticsContact">物流聯絡人</label>
                    <input type="text" id="editLogisticsContact" name="logisticsContact" maxlength="100">
                </div>
                <div class="form-group">
                    <label for="editMarketingContact">行銷聯絡人</label>
                    <input type="text" id="editMarketingContact" name="marketingContact" maxlength="100">
                </div>
                <div class="form-group">
                    <label for="editPhone">電話</label>
                    <input type="tel" id="editPhone" name="phone" maxlength="15">
                </div>
                -->
                <div class="form-group">
                    <label for="editAddress">地址</label>
                    <input type="text" id="editAddress" name="address" maxlength="100">
                </div>
                <!--
                <div class="form-group">
                    <label for="editEmail">電子郵件</label>
                    <input type="email" id="editEmail" name="email" maxlength="100">
                </div>
                -->
                <div class="form-group">
                    <label for="editWebsite">網址</label>
                    <input type="url" id="editWebsite" name="website" maxlength="250">
                </div>
                <div class="modal-actions">
                    <button type="submit" class="save-btn">儲存</button>
                    <button type="button" class="cancel-btn" onclick="hideEditVendorModal()">取消</button>
                </div>
            </form>
        </div>
    </div>

    <script>
        // Modal 控制
        function showAddVendorModal() {
            document.getElementById('vendorModal').style.display = 'flex';
            document.getElementById('vendorForm').reset();
            document.getElementById('vendorId').value = '';
        }

        function hideVendorModal() {
            document.getElementById('vendorModal').style.display = 'none';
        }

        function showEditVendorModal() {
            document.getElementById('editVendorModal').style.display = 'flex';
        }

        function hideEditVendorModal() {
            document.getElementById('editVendorModal').style.display = 'none';
        }

        // 廠商搜尋
        document.getElementById('vendorSearch').addEventListener('input', function(e) {
            const searchText = e.target.value.toLowerCase().trim();
            const rows = document.querySelectorAll('.vendors-table tbody tr');
            
            rows.forEach(row => {
                const code = row.cells[0].textContent.toLowerCase().replace('-', ''); // 母代號+子代號
                const uniformNumber = row.cells[1].textContent.toLowerCase(); // 統一編號
                const vendorName = row.cells[2].textContent.toLowerCase(); // 廠商名稱
                
                // 檢查是否符合任一搜尋條件
                const matchCode = code.includes(searchText.replace('-', '')); // 移除連字符進行比對
                const matchUniformNumber = uniformNumber.includes(searchText);
                const matchVendorName = vendorName.includes(searchText);
                
                // 如果符合任一條件就顯示該列
                row.style.display = (matchCode || matchUniformNumber || matchVendorName) ? '' : 'none';
            });
        });

        // 刪除廠商
        function deleteVendor(vendorId) {
            if (confirm('確定要刪除此廠商嗎？')) {
                fetch('delete_vendor.asp', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                    },
                    body: `vendorId=${vendorId}`
                })
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        location.reload();
                    } else {
                        alert(data.message);
                    }
                });
            }
        }

        // 啟用廠商
        function activateVendor(vendorId) {
            if (confirm('確定要啟用此廠商嗎？')) {
                fetch('activate_vendor.asp', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                    },
                    body: `vendorId=${vendorId}`
                })
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        location.reload();
                    } else {
                        alert(data.message);
                    }
                });
            }
        }

        // 編輯廠商
        function editVendor(vendorId) {
            fetch(`get_vendor.asp?id=${vendorId}`)
                .then(response => response.text())  // 先取得原始回應文字
                .then(text => {
                    try {
                        const data = JSON.parse(text);  // 嘗試解析 JSON
                        if (data.success) {
                            document.getElementById('editVendorId').value = data.VendorID;
                            document.getElementById('editParentCode').value = data.ParentCode;
                            document.getElementById('editChildCode').value = data.ChildCode;
                            document.getElementById('editUniformNumber').value = data.UniformNumber;
                            document.getElementById('editVendorName').value = data.VendorName;
                            document.getElementById('editContactPerson').value = data.ContactPerson;
                            document.getElementById('editLogisticsContact').value = data.LogisticsContact || '';
                            document.getElementById('editMarketingContact').value = data.MarketingContact || '';
                            document.getElementById('editPhone').value = data.Phone || '';
                            document.getElementById('editAddress').value = data.Address || '';
                            document.getElementById('editEmail').value = data.Email || '';
                            document.getElementById('editWebsite').value = data.Website || '';
                            
                            showEditVendorModal();
                        } else {
                            alert(data.message || '載入廠商資料失敗');
                        }
                    } catch (e) {
                        console.error('JSON 解析錯誤:', e);
                        console.error('收到的回應:', text);
                        alert('資料格式錯誤，請聯絡系統管理員');
                    }
                })
                .catch(error => {
                    console.error('網路錯誤:', error);
                    alert('載入廠商資料時發生錯誤，請稍後再試');
                });
        }
    </script>
</body>
</html> 