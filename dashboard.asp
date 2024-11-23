<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If
%>
<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>系統管理後台</title>
    <link rel="stylesheet" href="styles/dashboard.css">
</head>
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <!-- 主要內容區 -->
        <main class="main-content">
            <header class="top-bar">
                <div class="search-bar">
                    <input type="search" placeholder="搜尋...">
                </div>
                <div class="user-actions">                    
                    <span class="notification">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path d="M12 22C13.1 22 14 21.1 14 20H10C10 21.1 10.9 22 12 22ZM18 16V11C18 7.93 16.37 5.36 13.5 4.68V4C13.5 3.17 12.83 2.5 12 2.5C11.17 2.5 10.5 3.17 10.5 4V4.68C7.64 5.36 6 7.92 6 11V16L4 18V19H20V18L18 16ZM16 17H8V11C8 8.52 9.51 6.5 12 6.5C14.49 6.5 16 8.52 16 11V17Z" fill="currentColor"/>
                        </svg>
                    </span>
                    <span class="user-profile">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path d="M12 12C14.21 12 16 10.21 16 8C16 5.79 14.21 4 12 4C9.79 4 8 5.79 8 8C8 10.21 9.79 12 12 12ZM12 14C9.33 14 4 15.34 4 18V20H20V18C20 15.34 14.67 14 12 14Z" fill="currentColor"/>
                        </svg>
                    </span>
                </div>
            </header>

            <div class="content">
                <h1>儀表板</h1>
                <div class="stats-grid">
                    <div class="stat-card clickable" onclick="window.location.href='visit_questions.asp'">
                        <h3>訪廠題庫列表</h3>
                        <p class="number">25</p>
                    </div>
                    <div class="stat-card">
                        <h3>待訪廠數</h3>
                        <p class="number">8</p>
                    </div>
                    <div class="stat-card">
                        <h3>本月訪廠數</h3>
                        <p class="number">342</p>
                    </div>
                    <div class="stat-card">
                        <h3>廠商列表</h3>
                        <p class="number">56</p>
                    </div>
                </div>

                <div class="recent-activities">
                    <h2>最近活動</h2>
                    <div class="activity-list">
                        <!-- 活動列表將通過後端資料動態生成 -->
                    </div>
                </div>
            </div>
        </main>
    </div>

    <!-- 主題切換 JavaScript -->
    <script>
        // 檢查本地存儲中的主題設置
        const currentTheme = localStorage.getItem('theme') || 'light';
        document.documentElement.setAttribute('data-theme', currentTheme);
        document.getElementById('theme-toggle').checked = currentTheme === 'dark';

        // 主題切換監聽器
        document.getElementById('theme-toggle').addEventListener('change', function(e) {
            if(e.target.checked) {
                document.documentElement.setAttribute('data-theme', 'dark');
                localStorage.setItem('theme', 'dark');
            } else {
                document.documentElement.setAttribute('data-theme', 'light');
                localStorage.setItem('theme', 'light');
            }
        });
    </script>
</body>
</html> 