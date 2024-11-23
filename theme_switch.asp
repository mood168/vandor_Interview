<!-- 主題切換開關 -->
<div class="theme-switch-container">
<br/>
<span class="theme-text">黑白主題</span>
    <label class="theme-switch">
        <input type="checkbox" id="theme-toggle">        
        <div class="switch-wrapper">            
            <img class="theme-switch-icon" src="images/day-night.png" width="50" height="50" alt="日/夜切換圖示" class="theme-icon">
        </div>
    </label>
</div>

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