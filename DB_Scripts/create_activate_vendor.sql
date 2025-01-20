-- 建立啟用電商的預存程序
CREATE PROCEDURE sp_ActivateVendor
    @VendorID int,
    @ModifiedBy int
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRY
        BEGIN TRANSACTION
            
            UPDATE Vendors 
            SET IsActive = 1,
                ModifiedDate = GETDATE()
            WHERE VendorID = @VendorID
            
        COMMIT TRANSACTION
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION
            
        -- 拋出錯誤
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE()
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY()
        DECLARE @ErrorState INT = ERROR_STATE()
        
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END

