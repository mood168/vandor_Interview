<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 取得電商列表
Dim rsVendors
Set rsVendors = conn.Execute("SELECT VendorID, ParentCode, ChildCode, UniformNumber, VendorName, ContactPerson, Phone, Address, IsActive, CreatedDate FROM Vendors ORDER BY IsActive DESC, CreatedDate DESC")
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>電商訪談電商管理</title>
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
                    <input type="search" 
                           id="parentCodeSearch" 
                           placeholder="母代號..."
                           title="輸入母代號搜尋" 
                           class="save-btn" 
                           style="width: 140px;">
                           
                    <input type="search" 
                           id="childCodeSearch" 
                           placeholder="子代號..."
                           title="輸入子代號搜尋" 
                           class="save-btn" 
                           style="width: 140px;">
                           
                    <input type="search" 
                           id="uniformNumberSearch" 
                           placeholder="統一編號..."
                           title="輸入統一編號搜尋" 
                           class="save-btn" 
                           style="width: 180px;">
                           
                    <input type="search" 
                           id="vendorNameSearch" 
                           placeholder="電商名稱..."
                           title="輸入電商名稱搜尋" 
                           class="save-btn" 
                           style="width: 250px;">
                           
                    <input type="search" 
                           id="contactPersonSearch" 
                           placeholder="聯絡人..."
                           title="輸入聯絡人搜尋" 
                           class="save-btn" 
                           style="width: 160px;">
                </div>
                <div class="user-actions">
                    <button class="add-vendor-btn" onclick="showAddVendorModal()">新增電商</button>
                </div>
            </header>

            <div class="content">
                <h1>電商資料管理</h1>
                
                <div class="vendors-table-container">
                    <table class="vendors-table">
                        <thead>
                            <tr>
                                <th>母代號</th>
                                <th>子代號</th>
                                <th>統一編號</th>
                                <th>電商名稱</th>
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
                                    <td><%
                                        Dim uniformNum: uniformNum = rsVendors("UniformNumber")
                                        If Len(uniformNum) > 4 Then
                                            Response.Write Left(uniformNum,4) & String(Len(uniformNum)-4,"*")
                                        Else
                                            Response.Write uniformNum
                                        End If
                                    %></td>
                                    <td><%
                                        Dim vendorName: vendorName = rsVendors("VendorName")
                                        If Len(vendorName) > 1 Then
                                            If Asc(Left(vendorName,1)) > 255 Then
                                                ' Chinese characters
                                                Response.Write Left(vendorName,1) & "*" & Right(vendorName,Len(vendorName)-2)
                                            Else
                                                ' English characters
                                                If Len(vendorName) > 4 Then
                                                    Response.Write Left(vendorName,4) & String(Len(vendorName)-4,"*") 
                                                Else
                                                    Response.Write vendorName
                                                End If
                                            End If
                                        Else
                                            Response.Write vendorName
                                        End If
                                    %></td>
                                    <td><%
                                        Dim contactPerson: contactPerson = rsVendors("ContactPerson") 
                                        If Asc(Left(contactPerson,1)) > 255 Then
                                            ' Chinese characters
                                            Response.Write Left(contactPerson,1) & "*" & Right(contactPerson,Len(contactPerson)-2)
                                        Else
                                            ' English characters
                                            If Len(contactPerson) > 2 Then
                                                Response.Write Left(contactPerson,2) & String(Len(contactPerson)-2,"*")
                                            Else
                                                Response.Write contactPerson
                                            End If
                                        End If
                                    %></td>
                                    <td><%
                                        Dim phone: phone = rsVendors("Phone")
                                        If Len(phone) > 4 Then
                                            Response.Write Left(phone,4) & String(Len(phone)-4,"*")
                                        Else
                                            Response.Write phone
                                        End If
                                    %></td>
                                    <td><%
                                        Dim addr: addr = rsVendors("Address")
                                        If Len(addr) > 1 Then
                                            If Asc(Left(addr,1)) > 255 Then
                                                ' Chinese characters
                                                Response.Write Left(addr,1) & "*" & Right(addr,Len(addr)-2)
                                            Else
                                                ' English characters
                                                If Len(addr) > 4 Then
                                                    Response.Write Left(addr,4) & String(Len(addr)-4,"*")
                                                Else
                                                    Response.Write addr
                                                End If
                                            End If
                                        Else
                                            Response.Write addr
                                        End If
                                    %></td>
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

    <!-- 新增電商 Modal -->
    <div id="vendorModal" class="modal">
        <div class="modal-content">
            <h2>新增電商</h2>
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
                    <label for="vendorName">電商名稱</label>
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

    <!-- 編輯電商 Modal -->
    <div id="editVendorModal" class="modal">
        <div class="modal-content">
            <h2>編輯電商</h2>
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
                    <label for="editVendorName">電商名稱</label>
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

        // 電商搜尋
        function updateSearch() {
            const parentCodeText = document.getElementById('parentCodeSearch').value.toLowerCase().trim();
            const childCodeText = document.getElementById('childCodeSearch').value.toLowerCase().trim();
            const uniformNumberText = document.getElementById('uniformNumberSearch').value.toLowerCase().trim();
            const vendorNameText = document.getElementById('vendorNameSearch').value.toLowerCase().trim();
            const contactPersonText = document.getElementById('contactPersonSearch').value.toLowerCase().trim();
            
            const rows = document.querySelectorAll('.vendors-table tbody tr');
            
            rows.forEach(row => {
                const parentCode = row.cells[0].textContent.toLowerCase(); // 母代號
                const childCode = row.cells[1].textContent.toLowerCase(); // 子代號
                const uniformNumber = row.cells[2].textContent.toLowerCase(); // 統一編號
                const vendorName = row.cells[3].textContent.toLowerCase(); // 電商名稱
                const contactPerson = row.cells[4].textContent.toLowerCase(); // 聯絡人
                
                // 檢查是否符合所有搜尋條件
                const matchParentCode = !parentCodeText || parentCode.includes(parentCodeText);
                const matchChildCode = !childCodeText || childCode.includes(childCodeText);
                const matchUniformNumber = !uniformNumberText || uniformNumber.includes(uniformNumberText);
                const matchVendorName = !vendorNameText || vendorName.includes(vendorNameText);
                const matchContactPerson = !contactPersonText || contactPerson.includes(contactPersonText);
                
                // 所有條件都必須符合
                row.style.display = (matchParentCode && matchChildCode && matchUniformNumber && 
                                   matchVendorName && matchContactPerson) ? '' : 'none';
            });
        }

        // 為每個搜尋欄位添加事件監聽器
        document.getElementById('parentCodeSearch').addEventListener('input', updateSearch);
        document.getElementById('childCodeSearch').addEventListener('input', updateSearch);
        document.getElementById('uniformNumberSearch').addEventListener('input', updateSearch);
        document.getElementById('vendorNameSearch').addEventListener('input', updateSearch);
        document.getElementById('contactPersonSearch').addEventListener('input', updateSearch);

        // 刪除電商
        function deleteVendor(vendorId) {
            if (confirm('確定要刪除此電商嗎？')) {
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

        // 啟用電商
        function activateVendor(vendorId) {
            if (confirm('確定要啟用此電商嗎？')) {
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

        // 編輯電商
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
                            alert(data.message || '載入電商資料失敗');
                        }
                    } catch (e) {
                        console.error('JSON 解析錯誤:', e);
                        console.error('收到的回應:', text);
                        alert('資料格式錯誤，請聯絡系統管理員');
                    }
                })
                .catch(error => {
                    console.error('網路錯誤:', error);
                    alert('載入電商資料時發生錯誤，請稍後再試');
                });
        }
    </script>
</body>
</html> 