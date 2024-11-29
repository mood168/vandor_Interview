<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
' 檢查 token
token = Request.QueryString("token")
If token = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 檢查 token 是否有效
sql = "SELECT pr.UserID, u.FullName " & _
      "FROM PasswordResets pr " & _
      "INNER JOIN Users u ON pr.UserID = u.UserID " & _
      "WHERE pr.ResetToken = '" & Replace(token, "'", "''") & "' " & _
      "AND pr.ExpiryDate > GETDATE() " & _
      "AND pr.IsUsed = 0"

Set rs = conn.Execute(sql)

If rs.EOF Then
    Response.Redirect "login.html?error=" & Server.URLEncode("重設密碼連結已失效")
    Response.End
End If

userId = rs("UserID")
fullName = rs("FullName")
rs.Close
%>

<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>重設密碼</title>
    <link rel="stylesheet" href="styles/login.css">
    <link rel="stylesheet" href="styles/theme.css">
</head>
<body>
    <div class="container">
        <div class="reset-password-box">
            <h1>重設密碼</h1>
            <form id="resetPasswordForm" onsubmit="return submitResetPassword(event)">
                <input type="hidden" id="token" value="<%=token%>">
                <div class="input-group">
                    <input type="password" id="newPassword" required>
                    <label for="newPassword">新密碼</label>
                </div>
                <div class="input-group">
                    <input type="password" id="confirmPassword" required>
                    <label for="confirmPassword">確認新密碼</label>
                </div>
                <button type="submit" class="submit-button">確認重設</button>
            </form>
        </div>
    </div>

    <script>
    function submitResetPassword(event) {
        event.preventDefault();
        
        const newPassword = document.getElementById('newPassword').value;
        const confirmPassword = document.getElementById('confirmPassword').value;
        const token = document.getElementById('token').value;
        
        if (newPassword !== confirmPassword) {
            alert('兩次輸入的密碼不相符');
            return false;
        }
        
        fetch('update_password.asp', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: `token=${encodeURIComponent(token)}&password=${encodeURIComponent(newPassword)}`
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                alert('密碼已重設成功');
                window.location.href = 'login.html';
            } else {
                alert(data.message || '重設密碼失敗');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('發生錯誤，請稍後再試');
        });
        
        return false;
    }
    </script>
</body>
</html> 