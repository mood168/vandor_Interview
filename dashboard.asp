<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Redirect "login.html"
    Response.End
End If

' 處理搜尋請求
Dim searchKeyword, sql, rs
If Request.QueryString("searchKeyword") <> "" Then
    searchKeyword = Request.QueryString("searchKeyword")
Else
    searchKeyword = Request.Form("searchKeyword")
End If

If searchKeyword <> "" Then
    ' 建立模糊搜尋 SQL
    sql = "SELECT parentCode, ChildCode, VendorName, UniformNumber FROM Vendors " & _
          "WHERE parentCode LIKE '%" & searchKeyword & "%' " & _
          "OR ChildCode LIKE '%" & searchKeyword & "%' " & _
          "OR VendorName LIKE '%" & searchKeyword & "%' " & _
          "OR UniformNumber LIKE '%" & searchKeyword & "%' " & _
          "ORDER BY parentCode, ChildCode"
    
    Set rs = conn.Execute(sql)
End If
%>
<!DOCTYPE html>
<html lang="zh-TW" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>系統管理後台</title>
    <link rel="stylesheet" href="styles/dashboard.css">
    <style>
        .search-container {
            max-width: 1000px;
            margin: 50px auto;
            padding: 20px;
        }
        .search-form {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
        }
        .search-input {
            flex: 1;
            padding: 10px;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            background-color: var(--input-bg);
            color: var(--text-color);
        }
        .search-button {
            padding: 12px 24px;
            border-radius: 8px;
            font-size: 18px;
            cursor: pointer;
            transition: opacity 0.3s;
        }

        .search-button {
            background-color: transparent;
            border: 1px solid var(--border-color);
            color: var(--text-primary);
        }
        .search-results {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            background-color: var(--card-bg);
            color: var(--text-color);
        }
        .search-results th, 
        .search-results td {
            padding: 10px;
            border: 1px solid var(--border-color);
            text-align: left;
        }
        .search-results th {
            background-color: var(--header-bg);
            color: var(--text-color);
        }
        .search-results tr:hover {
            background-color: var(--hover-bg);
        }
        
        /* 深色模式特定樣式 */
        [data-theme="dark"] .search-results {
            border-color: var(--border-color);
        }
        [data-theme="dark"] .search-input {
            background-color: var(--input-bg);
            color: var(--text-color);
        }
        [data-theme="dark"] .search-button:hover {
            background-color: var(--primary-hover);
        }

        /* 新增廠商表單樣式 */
        .add-vendor-form {
            display: none; /* 預設隱藏表單 */
            margin-top: 30px;
            padding: 20px;
            border-radius: 8px;
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            max-width: 1000px;
            margin-left: auto;
            margin-right: auto;
        }
        .add-vendor-form h2 {
            margin-bottom: 20px;
            color: var(--text-color);
        }
        .form-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 20px;
        }
        .form-field {
            margin-bottom: 15px;
        }
        .form-field label {
            display: block;
            margin-bottom: 5px;
            color: var(--text-color);
        }
        .form-field input {
            width: 100%;
            padding: 8px;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            background-color: var(--input-bg);
            color: var(--text-color);
        }
        .form-buttons {
            margin-top: 20px;
            display: flex;
            gap: 10px;
            justify-content: flex-end;
        }
        .form-button {
            padding: 8px 20px;
            border-radius: 4px;
            cursor: pointer;
            border: 1px solid var(--border-color);
            background-color: var(--button-bg);
            color: var(--text-color);
        }
        .form-button.primary {
            background-color: var(--primary-color);
            color: white;
            border: none;
        }
    </style>
    <!-- 加入 jQuery 函式庫 -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    
    <script>
        $(document).ready(function() {
            let searchTimer;
            
            // 監聽搜尋框輸入事件
            $('.search-input').on('input', function() {
                clearTimeout(searchTimer);
                const searchValue = $(this).val();
                
                searchTimer = setTimeout(function() {
                    performSearch(searchValue);
                }, 300);
            });
            
            // 執行搜尋
            function performSearch(keyword) {
                $.ajax({
                    url: 'search_vendors.asp',
                    method: 'GET',
                    data: { searchKeyword: keyword },
                    success: function(response) {
                        $('.search-results-container').html(response);
                        // 為新加載的表格行加入點擊事件
                        bindTableRowClick();
                        
                        // 檢查是否有搜尋結果
                        if ($('.search-results tbody tr').length === 0 && keyword.trim() !== '') {
                            $('.add-vendor-form').show();
                        } else {
                            $('.add-vendor-form').hide();
                        }
                    },
                    error: function(xhr, status, error) {
                        console.error('搜尋發生錯誤:', error);
                    }
                });
            }

            // 綁定表格行點擊事件
            function bindTableRowClick() {
                $('.search-results tbody tr').css('cursor', 'pointer').on('click', function() {
                    const vendorName = $(this).find('td:eq(2)').text(); // 第三欄是廠商名稱
                    window.location.href = 'visit_questions.asp?vendor=' + encodeURIComponent(vendorName);
                });
            }

            // 初始綁定表格行點擊事件
            bindTableRowClick();
        });
    </script>
</head>
<body>
    <div class="dashboard-container">
        <!-- 側邊選單 -->
        <!--#include file="aside_menu.asp"-->

        <!-- 主要內容區 -->
        <main class="main-content">
            <div class="search-container">
                <div class="search-form">
                    <input type="text" name="searchKeyword" class="search-input" 
                           placeholder="請輸入 母代號 / 子代號 / 廠商名稱 / 統一編號 搜尋..." 
                           value="<%= searchKeyword %>">
                    <button type="button" class="search-button" onclick="window.location.href='dashboard.asp'">清除</button>                    
                </div>

                <!-- 搜尋結果容器 -->
                <div class="search-results-container">
                    <!-- 由 search_vendors.asp 控制內容 -->
                </div>
            </div>
            
        </main>
    </div>
</body>
</html>
<% 
If IsObject(rs) Then
    rs.Close
    Set rs = Nothing
End If
%> 