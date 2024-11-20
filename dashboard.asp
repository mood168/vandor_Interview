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
                    <!-- 主題切換開關 -->
                    <label class="theme-switch">
                        <input type="checkbox" id="theme-toggle">
                        <span class="slider"></span>
                    </label>
                    <span class="notification">🔔</span>
                    <span class="user-profile">👤</span>
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