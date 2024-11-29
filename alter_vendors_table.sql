-- 新增物流聯絡人和行銷聯絡人欄位
ALTER TABLE [dbo].[Vendors]
ADD [LogisticsContact] [nvarchar](100) NULL,
    [MarketingContact] [nvarchar](100) NULL
GO 