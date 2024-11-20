<!-- 主題切換開關 -->
<label class="theme-switch">
    <input type="checkbox" id="theme-toggle">
    <span class="slider"></span>
</label>

<script>
    // 檢查本地存儲中的主題設置
    const currentTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', currentTheme);
    
    // 設置切換開關的狀態
    const themeToggle = document.getElementById('theme-toggle');
    if (themeToggle) {
        themeToggle.checked = currentTheme === 'dark';
        
        // 主題切換監聽器
        themeToggle.addEventListener('change', function(e) {
            if(e.target.checked) {
                document.documentElement.setAttribute('data-theme', 'dark');
                localStorage.setItem('theme', 'dark');
            } else {
                document.documentElement.setAttribute('data-theme', 'light');
                localStorage.setItem('theme', 'light');
            }
        });
    }
</script> 