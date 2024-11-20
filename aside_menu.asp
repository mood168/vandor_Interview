<aside class="sidebar">
    <div class="logo">
        <h2>訪廠管理系統</h2>
    </div>
    <div class="user-info">
        <p>歡迎, <%=Session("FullName")%></p>
        <small><%=Session("UserRole")%></small>
    </div>
    <nav class="menu">
        <ul>
            <li><a href="dashboard.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/dashboard.asp" Then
                Response.Write("active")
            End If
            %>">
                <span class="icon">📊</span>儀表板
            </a></li>
            <li><a href="user_management.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/user_management.asp" Then
                Response.Write("active")
            End If
            %>">
                <span class="icon">👥</span>使用者管理
            </a></li>
            <li><a href="vendors_management.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/vendors_management.asp" Then
                Response.Write("active")
            End If
            %>">
                <span class="icon">🏢</span>廠商管理
            </a></li>
            <li><a href="#"><span class="icon">📅</span>訪廠預約</a></li>
            <li><a href="visit_questions.asp" class="<%
            If Request.ServerVariables("SCRIPT_NAME") = "/visit_questions.asp" Then
                Response.Write("active")
            End If
            %>">
                <span class="icon">📝</span>訪廠記錄
            </a></li>
            <li><a href="#"><span class="icon">⚙️</span>系統設定</a></li>
            <li><a href="logout.asp"><span class="icon">🚪</span>登出</a></li>
        </ul>
    </nav>
</aside>