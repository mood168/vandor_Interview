-- 開始交易
BEGIN TRANSACTION;

-- 新增雜湊密碼欄位（如果不存在）
IF NOT EXISTS (
    SELECT 1 
    FROM sys.columns 
    WHERE object_id = OBJECT_ID('Users') 
    AND name = 'PasswordHash'
)
BEGIN
    ALTER TABLE Users
    ADD PasswordHash NVARCHAR(64);
END
GO

-- 新增登入調試日誌表（如果不存在）
IF NOT EXISTS (
    SELECT 1 
    FROM sys.tables 
    WHERE name = 'LoginDebugLogs'
)
BEGIN
    CREATE TABLE LoginDebugLogs (
        ID INT IDENTITY(1,1) PRIMARY KEY,
        Username NVARCHAR(500),
        PasswordHash NVARCHAR(64),
        DebugMessage NVARCHAR(500),
        LoginTime DATETIME
    );
END
GO

-- 更新登入預存程序
CREATE OR ALTER PROCEDURE sp_UserLogin
    @Username NVARCHAR(500),
    @Password NVARCHAR(500),  -- 這裡接收的是已經雜湊過的密碼
    @LoginIP NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @UserID INT
    DECLARE @LoginStatus BIT = 0
    DECLARE @LoginMessage NVARCHAR(500) = N'登入失敗'
    
    -- 調試資訊
    INSERT INTO LoginDebugLogs (Username, PasswordHash, LoginTime)
    VALUES (@Username, @Password, GETDATE())
    
    -- 使用雜湊密碼進行驗證
    SELECT @UserID = UserID
    FROM Users
    WHERE Username = @Username 
    AND PasswordHash = @Password
    
    -- 調試資訊
    IF @UserID IS NULL
    BEGIN
        INSERT INTO LoginDebugLogs (Username, DebugMessage, LoginTime)
        VALUES (@Username, '未找到匹配的用戶記錄', GETDATE())
        
        -- 檢查用戶是否存在
        IF EXISTS (SELECT 1 FROM Users WHERE Username = @Username)
        BEGIN
            INSERT INTO LoginDebugLogs (Username, DebugMessage, LoginTime)
            VALUES (@Username, '用戶存在，但密碼不匹配', GETDATE())
            
            -- 獲取存儲的密碼雜湊
            DECLARE @StoredHash NVARCHAR(64)
            SELECT @StoredHash = PasswordHash
            FROM Users
            WHERE Username = @Username
            
            INSERT INTO LoginDebugLogs (Username, DebugMessage, LoginTime)
            VALUES (@Username, '存儲的密碼雜湊: ' + ISNULL(@StoredHash, 'NULL'), GETDATE())
        END
        ELSE
        BEGIN
            INSERT INTO LoginDebugLogs (Username, DebugMessage, LoginTime)
            VALUES (@Username, '用戶名不存在', GETDATE())
        END
    END
    ELSE
    BEGIN
        SET @LoginStatus = 1
        SET @LoginMessage = N'登入成功'
        
        -- 記錄登入資訊
        INSERT INTO LoginLogs (UserID, LoginIP, LoginTime, Status)
        VALUES (@UserID, @LoginIP, GETDATE(), 1)
    END
    
    -- 返回登入結果
    SELECT 
        u.UserID,
        u.Username,
        u.FullName,
        u.UserRole,
        @LoginStatus as LoginStatus,
        @LoginMessage as LoginMessage
    FROM Users u
    WHERE u.UserID = @UserID
END
GO

-- 建立密碼更新預存程序
CREATE OR ALTER PROCEDURE sp_UpdateUserPassword
    @UserID INT,
    @NewPasswordHash NVARCHAR(64)
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE Users
    SET PasswordHash = @NewPasswordHash
    WHERE UserID = @UserID
    
    RETURN 0
END
GO

-- 為所有使用者設定臨時密碼
-- 臨時密碼為：P@ssw0rd123
-- 對應的 SHA-256 雜湊值（由 JavaScript 生成）
UPDATE Users
SET PasswordHash = '8A9BCF1E6A5D92EC6877B0D76DD51ED00B2675D08B0BE48393F93784D2BF0A5C'
WHERE PasswordHash IS NULL;

-- 如果更新成功，提交交易
IF @@ERROR = 0
BEGIN
    COMMIT TRANSACTION;
    PRINT N'資料庫更新成功！';
    PRINT N'所有使用者的臨時密碼已設定為：P@ssw0rd123';
END
ELSE
BEGIN
    ROLLBACK TRANSACTION;
    PRINT N'發生錯誤，資料庫更新已回滾。';
END
