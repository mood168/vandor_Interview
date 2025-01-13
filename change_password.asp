<%@ Language="VBScript" CodePage="65001" %>
<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>變更密碼 | 系統管理</title>
    <link rel="stylesheet" href="styles/theme.css">
    <style>
        .container {
            max-width: 500px;
            margin: 50px auto;
            padding: 20px;
        }
        .form-group {
            margin-bottom: 15px;
        }
        .form-group label {
            display: block;
            margin-bottom: 5px;
        }
        .form-group input {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        .error-message {
            color: red;
            margin-bottom: 10px;
        }
        .password-rules {
            margin: 15px 0;
            padding: 10px;
            background: #f8f8f8;
            border-radius: 4px;
        }
        .submit-button {
            background: #007bff;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .submit-button:hover {
            background: #0056b3;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>變更密碼</h1>
        
        <% If Request.QueryString("expired") = "1" Then %>
            <div class="error-message">您的密碼已過期，請設定新密碼。</div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
            <div class="error-message"><%=Server.HTMLEncode(Request.QueryString("error"))%></div>
        <% End If %>

        <div class="password-rules">
            <h3>密碼規則：</h3>
            <ul>
                <li>至少6個字符</li>
                <li>必須包含大寫字母</li>
                <li>必須包含小寫字母</li>
                <li>必須包含數字</li>
                <li>每3個月需要更換一次密碼</li>
            </ul>
        </div>

        <form method="post" action="change_password_process.asp" onsubmit="return validateForm()">
            <div class="form-group">
                <label for="currentPassword">目前密碼：</label>
                <input type="password" id="currentPassword" name="currentPassword" required>
            </div>

            <div class="form-group">
                <label for="newPassword">新密碼：</label>
                <input type="password" id="newPassword" name="newPassword" required>
            </div>

            <div class="form-group">
                <label for="confirmPassword">確認新密碼：</label>
                <input type="password" id="confirmPassword" name="confirmPassword" required>
            </div>

            <button type="submit" class="submit-button">變更密碼</button>
        </form>
    </div>

    <script>
        function validateForm() {
            var newPassword = document.getElementById('newPassword').value;
            var confirmPassword = document.getElementById('confirmPassword').value;
            
            // 檢查密碼規則
            var passwordRegex = /^(?=.*[a-z])(?=.*[A-Z])(?=.*\d).{6,}$/;
            if (!passwordRegex.test(newPassword)) {
                alert('新密碼不符合密碼規則要求');
                return false;
            }
            
            // 檢查密碼確認
            if (newPassword !== confirmPassword) {
                alert('新密碼與確認密碼不相符');
                return false;
            }
            
            return true;
        }
    </script>
</body>
</html> 