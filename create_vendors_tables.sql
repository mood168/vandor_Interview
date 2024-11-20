-- 建立廠商資料表
CREATE TABLE [dbo].[Vendors](
    [VendorID] [int] IDENTITY(1,1) PRIMARY KEY,
    [ParentCode] [char](3) NOT NULL,
    [ChildCode] [char](3) NOT NULL,
    [UniformNumber] [char](8) NOT NULL,
    [VendorName] [nvarchar](100) NOT NULL,
    [ContactPerson] [nvarchar](100) NOT NULL,
    [Phone] [varchar](15) NULL,
    [Address] [nvarchar](100) NULL,
    [Email] [varchar](100) NULL,
    [Website] [varchar](250) NULL,
    [CreatedDate] [datetime] DEFAULT GETDATE(),
    [ModifiedDate] [datetime] DEFAULT GETDATE(),
    [IsActive] [bit] DEFAULT 1
)
GO

-- 建立唯一索引確保代號組合不重複
CREATE UNIQUE INDEX [IX_Vendors_Codes] ON [dbo].[Vendors]
(
    [ParentCode] ASC,
    [ChildCode] ASC
)
GO

-- 建立唯一索引確保統一編號不重複
CREATE UNIQUE INDEX [IX_Vendors_UniformNumber] ON [dbo].[Vendors]
(
    [UniformNumber] ASC
)
GO

-- 建立預存程序：新增廠商
CREATE PROCEDURE [dbo].[sp_AddVendor]
    @ParentCode char(3),
    @ChildCode char(3),
    @UniformNumber char(8),
    @VendorName nvarchar(100),
    @ContactPerson nvarchar(100),
    @Phone varchar(15),
    @Address nvarchar(100),
    @Email varchar(100),
    @Website varchar(250),
    @CreatedBy int
AS
BEGIN
    SET NOCOUNT ON;
    
    IF EXISTS (SELECT 1 FROM [dbo].[Vendors] WHERE ParentCode = @ParentCode AND ChildCode = @ChildCode)
    BEGIN
        RAISERROR ('代號組合已存在', 16, 1)
        RETURN
    END
    
    IF EXISTS (SELECT 1 FROM [dbo].[Vendors] WHERE UniformNumber = @UniformNumber)
    BEGIN
        RAISERROR ('統一編號已存在', 16, 1)
        RETURN
    END
    
    INSERT INTO [dbo].[Vendors] (
        ParentCode, ChildCode, UniformNumber, VendorName, 
        ContactPerson, Phone, Address, Email, Website
    )
    VALUES (
        @ParentCode, @ChildCode, @UniformNumber, @VendorName,
        @ContactPerson, @Phone, @Address, @Email, @Website
    )
    
    SELECT SCOPE_IDENTITY() AS VendorID
END
GO

-- 建立預存程序：更新廠商
CREATE PROCEDURE [dbo].[sp_UpdateVendor]
    @VendorID int,
    @VendorName nvarchar(100),
    @ContactPerson nvarchar(100),
    @Phone varchar(15),
    @Address nvarchar(100),
    @Email varchar(100),
    @Website varchar(250),
    @ModifiedBy int
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE [dbo].[Vendors]
    SET VendorName = @VendorName,
        ContactPerson = @ContactPerson,
        Phone = @Phone,
        Address = @Address,
        Email = @Email,
        Website = @Website,
        ModifiedDate = GETDATE()
    WHERE VendorID = @VendorID
END
GO

-- 建立預存程序：刪除廠商（軟刪除）
CREATE PROCEDURE [dbo].[sp_DeleteVendor]
    @VendorID int,
    @ModifiedBy int
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE Vendors
    SET IsActive = 0,
        ModifiedDate = GETDATE()
    WHERE VendorID = @VendorID
END
GO 