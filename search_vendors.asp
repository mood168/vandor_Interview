<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.End
End If

Dim searchKeyword, sql, rs
searchKeyword = Trim(Request.QueryString("searchKeyword"))

' 只在有搜尋關鍵字時執行搜尋
If Len(searchKeyword) > 0 Then
    ' 建立模糊搜尋 SQL
    sql = "SELECT ParentCode, ChildCode, VendorName, UniformNumber FROM Vendors " & _
          "WHERE ParentCode LIKE '%" & searchKeyword & "%' " & _
          "OR ChildCode LIKE '%" & searchKeyword & "%' " & _
          "OR VendorName LIKE '%" & searchKeyword & "%' " & _
          "OR UniformNumber LIKE '%" & searchKeyword & "%' " & _
          "ORDER BY ParentCode, ChildCode"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
%>
        <table class="search-results">
            <thead>
                <tr>
                    <th>母代號</th>
                    <th>子代號</th>
                    <th>廠商名稱</th>
                    <th>統一編號</th>
                </tr>
            </thead>
            <tbody>
                <% Do While Not rs.EOF %>
                    <tr style="cursor: pointer;">
                        <td><%= rs("ParentCode") %></td>
                        <td><%= rs("ChildCode") %></td>
                        <td><%= rs("VendorName") %></td>
                        <td><%= rs("UniformNumber") %></td>
                    </tr>
                <%
                    rs.MoveNext
                Loop
                %>
            </tbody>
        </table>
<%
    Else
        ' 只有在有搜尋關鍵字且查無結果時才顯示表單
%>
        <div class="add-vendor-form">
            <h2>新增廠商資訊</h2>
            <form action="save_vendor.asp" method="post">
                <div class="form-grid">
                    <div class="form-field">
                        <label for="parentCode">母代號</label>
                        <input type="text" id="parentCode" name="parentCode" maxlength="3" required 
                               pattern="[0-9]{3}" title="請輸入3碼英數字" value="">
                    </div>
                    <div class="form-field">
                        <label for="childCode">子代號</label>
                        <input type="text" id="childCode" name="childCode" maxlength="3" required 
                               pattern="[0-9]{3}" title="請輸入3碼英數字">
                    </div>
                </div>
                <div class="form-field">
                    <label for="uniformNumber">統一編號</label>
                    <input type="text" id="uniformNumber" name="uniformNumber" maxlength="8" required
                           pattern="[0-9]{8}" title="請輸入8碼數字">
                </div>
                <div class="form-field">
                    <label for="vendorName">廠商名稱</label>
                    <input type="text" id="vendorName" name="vendorName" maxlength="100" required>
                </div>
                <div class="form-field">
                    <label for="contactPerson">聯絡人</label>
                    <input type="text" id="contactPerson" name="contactPerson" maxlength="100">
                </div>
                <div class="form-field">
                    <label for="phone">電話</label>
                    <input type="tel" id="phone" name="phone" maxlength="15">
                </div>
                <div class="form-field">
                    <label for="address">地址</label>
                    <input type="text" id="address" name="address" maxlength="100">
                </div>
                <div class="form-field">
                    <label for="email">電子郵件</label>
                    <input type="email" id="email" name="email" maxlength="100">
                </div>
                <div class="form-field">
                    <label for="website">網址</label>
                    <input type="url" id="website" name="website" maxlength="250">
                </div>
                <div class="form-buttons">
                    <button type="button" class="form-button" onclick="window.location.href='dashboard.asp'">取消</button>
                    <button type="submit" class="form-button primary">儲存</button>
                </div>
            </form>
        </div>
<%
    End If
    
    If IsObject(rs) Then
        rs.Close
        Set rs = Nothing
    End If
End If
%> 