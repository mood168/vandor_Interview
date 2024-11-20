-- 建立訪廠記錄主表
CREATE TABLE [dbo].[VisitRecords](
    [VisitID] [int] IDENTITY(1,1) PRIMARY KEY,
    [CompanyName] [nvarchar](200) NOT NULL,
    [VisitDate] [datetime] NOT NULL,
    [VisitorID] [int] FOREIGN KEY REFERENCES Users(UserID),
    [Status] [nvarchar](20) DEFAULT 'Draft',  -- Draft, Completed, Reviewed
    [CreatedDate] [datetime] DEFAULT GETDATE(),
    [ModifiedDate] [datetime] DEFAULT GETDATE()
)
GO

-- 建立訪廠答案表
CREATE TABLE [dbo].[VisitAnswers](
    [AnswerID] [int] IDENTITY(1,1) PRIMARY KEY,
    [VisitID] [int] FOREIGN KEY REFERENCES VisitRecords(VisitID),
    [QuestionID] [int] FOREIGN KEY REFERENCES VisitQuestions(QuestionID),
    [Answer] [nvarchar](max),
    [ModifiedDate] [datetime] DEFAULT GETDATE()
)
GO

-- 建立題庫分類表
CREATE TABLE [dbo].[QuestionCategories](
    [CategoryID] [int] IDENTITY(1,1) PRIMARY KEY,
    [CategoryName] [nvarchar](100) NOT NULL,
    [SortOrder] [int] DEFAULT 0,
    [IsRequired] [bit] DEFAULT 1
)
GO

-- 建立題庫表
CREATE TABLE [dbo].[VisitQuestions](
    [QuestionID] [int] IDENTITY(1,1) PRIMARY KEY,
    [CategoryID] [int] FOREIGN KEY REFERENCES QuestionCategories(CategoryID),
    [QuestionText] [nvarchar](500) NOT NULL,
    [IsRequired] [bit] DEFAULT 1,
    [CanModify] [bit] DEFAULT 1,
    [HasOptions] [bit] DEFAULT 0,
    [Options] [nvarchar](max), -- JSON格式存儲選項
    [SortOrder] [int] DEFAULT 0,
    [Status] [bit] DEFAULT 1
)
GO

-- 插入題庫分類
INSERT INTO QuestionCategories (CategoryName, SortOrder, IsRequired) VALUES 
(N'公司概況', 1, 1),
(N'營運狀況', 2, 1),
(N'跨國佈局', 3, 0),
(N'廣告採買', 4, 0),
(N'新服務', 5, 1),
(N'行銷配合', 6, 1),
(N'金流', 7, 1),
(N'物流', 8, 1)
GO

-- 插入題庫問題
INSERT INTO VisitQuestions 
(CategoryID, QuestionText, IsRequired, CanModify, HasOptions, Options, SortOrder, Status) VALUES 
(1, N'成立時間', 1, 1, 0, NULL, 1, 1),
(1, N'有網站嗎', 1, 1, 1, N'["有","無"]', 2, 1),
(1, N'系統商', 1, 1, 0, NULL, 3, 1),
(1, N'自有倉庫', 1, 1, 1, N'["有","無"]', 4, 1),
(1, N'公司人數', 1, 1, 0, NULL, 5, 1),
(1, N'辦公室或公司總坪數', 1, 1, 0, NULL, 6, 1),
(1, N'商品種類', 1, 1, 0, NULL, 7, 1),
(1, N'商品來源', 1, 1, 0, NULL, 8, 1),
(1, N'銷售通路', 1, 1, 0, NULL, 9, 1),
(2, N'每月出貨量', 1, 1, 0, NULL, 1, 1),
(2, N'均單價', 1, 1, 0, NULL, 2, 1),
(2, N'TA', 1, 1, 0, NULL, 3, 1),
(2, N'活動週期', 1, 1, 0, NULL, 4, 1),
(2, N'未來規劃', 1, 1, 0, NULL, 5, 1),
(3, N'目前跨境經驗', 1, 1, 0, NULL, 1, 1),
(3, N'訂單數', 1, 1, 0, NULL, 2, 1),
(3, N'經營困難', 1, 1, 0, NULL, 3, 1),
(3, N'國內外消費者需求差異', 1, 1, 0, NULL, 4, 1),
(3, N'展望', 1, 1, 0, NULL, 5, 1),
(4, N'廣告預算', 1, 1, 0, NULL, 1, 1),
(5, N'新服務', 1, 1, 0, NULL, 1, 1),
(6, N'可以導入SP嗎?', 1, 1, 1, N'["是","否"]', 1, 1),
(6, N'預購', 1, 1, 1, N'["是","否"]', 2, 1),
(6, N'LINE OA', 1, 1, 1, N'["是","否"]', 3, 1),
(7, N'付款方式', 1, 1, 1, NULL, 1, 1),
(8, N'宅配', 1, 1, 1, N'["是","否"]', 1),
(8, N'配到大智通', 1, 1, 1, NULL, 2, 1),
(8, N'超取通路', 1, 1, 1, NULL, 3, 1),
(8, N'新增通路', 1, 1, 1, NULL, 4, 1),
GO 