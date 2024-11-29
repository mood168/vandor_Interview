-- 建立使用者資料表
CREATE TABLE [dbo].[Users](
    [UserID] [int] IDENTITY(1,1) PRIMARY KEY,
    [Username] [nvarchar](500) NOT NULL UNIQUE,
    [Password] [nvarchar](500) NOT NULL,
    [FullName] [nvarchar](100) NOT NULL,
    [Phone] [nvarchar](100) NULL,
    [Email] [nvarchar](100) NULL,
    [Department] [nvarchar](50) NULL,
    [UserRole] [nvarchar](20) NOT NULL,  -- Admin, Manager, User
    [IsActive] [bit] DEFAULT 1,
    [LastLoginTime] [datetime] NULL,
    [CreatedDate] [datetime] DEFAULT GETDATE(),
    [ModifiedDate] [datetime] DEFAULT GETDATE()
)
GO

-- 建立登入記錄表
CREATE TABLE [dbo].[LoginLogs](
    [LogID] [int] IDENTITY(1,1) PRIMARY KEY,
    [UserID] [int] FOREIGN KEY REFERENCES Users(UserID),
    [LoginTime] [datetime] DEFAULT GETDATE(),
    [LoginIP] [varchar](50) NULL,
    [LoginStatus] [bit] DEFAULT 1,  -- 1:成功, 0:失敗
    [LoginMessage] [nvarchar](255) NULL
)
GO

-- 建立初始管理員帳號（密碼為 'admin123'）
INSERT INTO [dbo].[Users]
(Username, Password, FullName, Phone, Email, Department, UserRole)
VALUES
('admin', '@@MYjoan1391', N'系統管理員', '0909881391', 'tonymac168@gmail.com', N'開發者', 'Admin')
GO

-- 建立預存程序：使用者登入驗證
CREATE PROCEDURE [dbo].[usp_UserLogin]
    @Username nvarchar(50),
    @Password nvarchar(255),
    @LoginIP varchar(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @UserID int
    DECLARE @LoginStatus bit
    DECLARE @LoginMessage nvarchar(255)

    -- 檢查使用者帳號密碼
    SELECT @UserID = UserID
    FROM Users
    WHERE Username = @Username 
    AND Password = @Password
    AND IsActive = 1

    IF @UserID IS NOT NULL
    BEGIN
        SET @LoginStatus = 1
        SET @LoginMessage = N'登入成功'
        
        -- 更新最後登入時間
        UPDATE Users 
        SET LastLoginTime = GETDATE()
        WHERE UserID = @UserID
    END
    ELSE
    BEGIN
        SET @LoginStatus = 0
        SET @LoginMessage = N'帳號或密碼錯誤'
    END

    -- 記錄登入日誌
    INSERT INTO LoginLogs 
    (UserID, LoginIP, LoginStatus, LoginMessage)
    VALUES 
    (@UserID, @LoginIP, @LoginStatus, @LoginMessage)

    -- 回傳登入結果
    SELECT 
        @LoginStatus AS LoginStatus,
        @LoginMessage AS LoginMessage,
        u.UserID,
        u.Username,
        u.FullName,
        u.Email,
        u.Department,
        u.UserRole
    FROM Users u
    WHERE u.UserID = @UserID
END
GO 