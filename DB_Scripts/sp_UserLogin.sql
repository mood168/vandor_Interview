USE [Vendor_Interview_V2_Test]
GO
/****** Object:  StoredProcedure [dbo].[sp_UserLogin]    Script Date: 2025/1/18 18:24:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[sp_UserLogin]
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
        INSERT INTO LoginLogs (UserID, LoginTime, LoginIP, LoginStatus)
        VALUES (@UserID, GETDATE(), @LoginIP, 1)
        
        -- 回傳使用者資訊
        SELECT 
            u.UserID,
            u.Username,
            u.FullName,
            u.UserRole,
			u.IsActive,
            u.ModifiedDate,
            @LoginStatus AS LoginStatus,
            @LoginMessage AS LoginMessage
        FROM Users u
        WHERE u.UserID = @UserID
    END
    ELSE
    BEGIN
        -- 記錄失敗的登入嘗試
        INSERT INTO LoginLogs (UserID, LoginTime, LoginIP, LoginStatus)
        VALUES (NULL, GETDATE(), @LoginIP, 0)
        
        -- 回傳錯誤訊息
        SELECT 
            NULL AS UserID,
            NULL AS Username,
            NULL AS FullName,
            NULL AS UserRole,
			NULL AS IsActive,
            NULL AS ModifiedDate,
            @LoginStatus AS LoginStatus,
            @LoginMessage AS LoginMessage
    END
END
