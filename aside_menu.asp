<button id="menuToggle" class="menu-toggle">
    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M3 12H21" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        <path d="M3 6H21" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        <path d="M3 18H21" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
    </svg>
</button>

<aside class="sidebar" id="sidebar">
    <div>
        <h2>電商訪談系統</h2> 
        <br/>       
        <p class="user-info" style="background-color: var(--bg-primary);">
            歡迎, <%=Session("FullName")%> (<small><%=Session("UserRole")%></small>)
        </p>
    </div>
    
    <nav class="menu">
        <ul>
            <li><a href="dashboard.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/dashboard.asp" Then
                Response.Write("active")
            End If
            %>" style="font-size: 1.2rem;">
                <span class="icon">🏠</span>首頁 HOME
            </a></li>
            <li><a href="visit_questions.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/visit_questions.asp" Then
                Response.Write("active")
            End If
            %>" style="font-size: 1.2rem;">
                <span class="icon">📝</span>電商訪談填寫
            </a></li>
            <li><a href="visit_management.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/visit_management.asp" Then
                Response.Write("active")
            End If
            %>" style="font-size: 1.2rem;">
                <span class="icon">📅</span>電商訪談管理
            </a></li>
            <li><a href="visit_management_by_date.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/visit_management_by_date.asp" Then
                Response.Write("active")
            End If
            %>" style="font-size: 1.2rem;">
                <span class="icon">📅</span>電商訪談紀錄
            </a></li>
            <li><a href="vendors_management.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/vendors_management.asp" Then
                Response.Write("active")
            End If
            %>" style="font-size: 1.2rem;">
                <span class="icon">🏢</span>電商資料管理
            </a></li> 
            <li><a href="user_management.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/user_management.asp" Then
                Response.Write("active")
            End If
            %>" style="font-size: 1.2rem;">
                <span class="icon">👥</span>用戶資料管理
            </a></li>                     
            <li><a href="logout.asp" style="font-size: 1.2rem;"><span class="icon">🚪</span>登出</a></li>
        </ul>
    </nav>
    <div style="text-align: center;background-color: var(--bg-primary-light);">
        <hr><!--#include file="theme_switch.asp"-->
    </div>
</aside>

<script>
    // 選單切換功能
    const menuToggle = document.getElementById('menuToggle');
    const sidebar = document.getElementById('sidebar');
    const overlay = document.createElement('div');
    overlay.className = 'sidebar-overlay';
    document.body.appendChild(overlay);

    menuToggle.addEventListener('click', () => {
        sidebar.classList.toggle('active');
        overlay.classList.toggle('active');
        document.body.classList.toggle('menu-open');
    });

    // 點擊遮罩層關閉選單
    overlay.addEventListener('click', () => {
        sidebar.classList.remove('active');
        overlay.classList.remove('active');
        document.body.classList.remove('menu-open');
    });

    // 監聽視窗大小變化
    window.addEventListener('resize', () => {
        if (window.innerWidth > 770) {
            sidebar.classList.remove('active');
            overlay.classList.remove('active');
            document.body.classList.remove('menu-open');
        }
    });
</script>