-- 檢查並新增 LastPasswordChangeDate 欄位
IF NOT EXISTS (
    SELECT * FROM sys.columns 
    WHERE object_id = OBJECT_ID('Users') 
    AND name = 'LastPasswordChangeDate'
)
BEGIN
    ALTER TABLE Users ADD LastPasswordChangeDate DATETIME DEFAULT GETDATE()
END
GO

-- 更新現有記錄的 LastPasswordChangeDate
UPDATE Users 
SET LastPasswordChangeDate = GETDATE() 
WHERE LastPasswordChangeDate IS NULL
GO

-- 修改 sp_UserLogin 預存程序
IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'sp_UserLogin')
    DROP PROCEDURE sp_UserLogin
GO

CREATE PROCEDURE sp_UserLogin
    @Username NVARCHAR(500),
    @Password NVARCHAR(500),
    @LoginIP NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @UserID INT
    DECLARE @LoginStatus BIT = 0
    DECLARE @LoginMessage NVARCHAR(500) = N'登入失敗'
    
    -- 檢查使用者是否存在且密碼正確
    SELECT @UserID = UserID
    FROM Users
    WHERE Username = @Username 
    AND Password = @Password
    AND IsActive = 1
    
    IF @UserID IS NOT NULL
    BEGIN
        SET @LoginStatus = 1
        SET @LoginMessage = N'登入成功'
        
        -- 記錄登入歷史
        INSERT INTO LoginHistory (UserID, LoginTime, LoginIP, LoginStatus)
        VALUES (@UserID, GETDATE(), @LoginIP, 1)
        
        -- 回傳使用者資訊
        SELECT 
            u.UserID,
            u.Username,
            u.FullName,
            u.UserRole,
            u.LastPasswordChangeDate,
            @LoginStatus AS LoginStatus,
            @LoginMessage AS LoginMessage
        FROM Users u
        WHERE u.UserID = @UserID
    END
    ELSE
    BEGIN
        -- 記錄失敗的登入嘗試
        INSERT INTO LoginHistory (UserID, LoginTime, LoginIP, LoginStatus)
        VALUES (NULL, GETDATE(), @LoginIP, 0)
        
        -- 回傳錯誤訊息
        SELECT 
            NULL AS UserID,
            NULL AS Username,
            NULL AS FullName,
            NULL AS UserRole,
            NULL AS LastPasswordChangeDate,
            @LoginStatus AS LoginStatus,
            @LoginMessage AS LoginMessage
    END
END
GO

-- 建立 sp_CheckCurrentPassword 預存程序
IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'sp_CheckCurrentPassword')
    DROP PROCEDURE sp_CheckCurrentPassword
GO

CREATE PROCEDURE sp_CheckCurrentPassword
    @UserID INT,
    @CurrentPassword NVARCHAR(500)
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT CASE 
        WHEN EXISTS (
            SELECT 1 
            FROM Users 
            WHERE UserID = @UserID 
            AND Password = @CurrentPassword
            AND IsActive = 1
        ) THEN 1 
        ELSE 0 
    END AS IsValid
END
GO

-- 建立 sp_UpdatePassword 預存程序
IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'sp_UpdatePassword')
    DROP PROCEDURE sp_UpdatePassword
GO

CREATE PROCEDURE sp_UpdatePassword
    @UserID INT,
    @NewPassword NVARCHAR(500)
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE Users 
    SET Password = @NewPassword,
        LastPasswordChangeDate = GETDATE(),
        ModifiedDate = GETDATE()
    WHERE UserID = @UserID
    AND IsActive = 1
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS Success, N'密碼更新成功' AS Message
    ELSE
        SELECT 0 AS Success, N'密碼更新失敗' AS Message
END
GO 