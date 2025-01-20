-- 建立訪廠記錄主表
CREATE TABLE [dbo].[VisitRecords](
    [VisitID] [int] IDENTITY(1,1) PRIMARY KEY,
    [CompanyName] [nvarchar](200) NOT NULL,
    [VisitDate] [datetime] NOT NULL,
    [VisitorID] [int] FOREIGN KEY REFERENCES Users(UserID),
    [Status] [nvarchar](20) DEFAULT 'Draft',  -- Draft, Completed, Reviewed
    [CreatedDate] [datetime] DEFAULT GETDATE(),
    [ModifiedDate] [datetime] DEFAULT GETDATE(),
    [Interviewee] [nvarchar](100) NULL
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
    [AnswerType] [nvarchar](50) DEFAULT 'text', -- text, radio, checkbox, select, number, date
    [Options] [nvarchar](max), -- JSON格式存儲選項
    [SortOrder] [int] DEFAULT 0,
    [Status] [bit] DEFAULT 1,
    [HasPercentage] [bit] DEFAULT 0  -- 新增：標記是否需要輸入百分比
)
GO

-- 建立訪廠答案表
CREATE TABLE [dbo].[VisitAnswers](
    [AnswerID] [int] IDENTITY(1,1) PRIMARY KEY,
    [VisitID] [int] FOREIGN KEY REFERENCES VisitRecords(VisitID),
    [QuestionID] [int] FOREIGN KEY REFERENCES VisitQuestions(QuestionID),
    [Answer] [nvarchar](max) NULL,
    [ModifiedDate] [datetime] DEFAULT GETDATE()
)
GO


-- 插入題庫分類
INSERT INTO QuestionCategories (CategoryName, SortOrder, IsRequired) VALUES 
(N'公司概況', 1, 1),
(N'營運狀況', 2, 1),
(N'金流狀況', 3, 1),
(N'物流狀況', 4, 1),
(N'跨國佈局', 5, 1),
(N'廣告採買', 6, 1),
(N'行銷配合', 7, 1),
(N'新服務', 8, 1),
(N'經驗分享', 9, 1)  -- 新增分類
GO

-- 插入題庫問題
INSERT INTO VisitQuestions 
(CategoryID, QuestionText, IsRequired, CanModify, HasOptions, AnswerType, Options, SortOrder, HasPercentage) VALUES 
-- 公司概況
(1, N'公司成立時間 西元 ___ 年', 1, 1, 0, 'date', NULL, 1, 0),
(1, N'網站成立時間 西元 ___ 年', 1, 1, 0, 'date', NULL, 2, 0),
(1, N'網站系統商', 1, 1, 1, 'radio', N'["自建","委外"]', 3, 0),
(1, N'倉儲物流', 1, 1, 1, 'radio', N'["自有倉庫出貨","委外"]', 4, 0),
(1, N'公司人數 ___ 人', 1, 1, 0, 'number', NULL, 5, 0),
(1, N'辦公室 ___ 坪', 1, 1, 0, 'number', NULL, 6, 0),
(1, N'倉庫坪數 ___ 坪', 1, 1, 0, 'number', NULL, 7, 0),
(1, N'商品種類', 1, 1, 1, 'checkbox', N'["流行衣飾","包包配件","運動健身","戶外旅行","男女鞋","嬰幼童與母親","美妝保健","文創商品","書籍及雜誌期刊","居家生活","美食、伴手禮","寵物","家電3C","電玩娛樂、收藏","貨運倉儲"]', 8, 0),
(1, N'商品來源', 1, 1, 1, 'checkbox', N'["本地","大陸","日本","韓國","美國","歐洲","新馬","港澳","泰國","菲律賓","越南","加拿大","紐澳"]', 9, 0),
(1, N'銷售通路', 1, 1, 1, 'checkbox', N'["實體通路：自有門市佔%","百貨量販佔%","超市便利商店佔%","虛擬通路：自有官網佔%","蝦皮佔%","MOMO佔%","PCHOME佔%","YAHOO佔%","海外通路：線上佔%","經銷商(批發)佔%"]', 10, 1),

-- 營運狀況
(2, N'每月出貨量', 1, 1, 1, 'radio', N'["500件↓","501~1500件","1501~3000件","3001~5000件","5000件↑"]', 1, 0),
(2, N'每月均單價', 1, 1, 1, 'radio', N'["500元↓","501~1000元","1001~1500元","1501~2000元","2000元↑"]', 2, 0),
(2, N'TA輪廓', 1, 1, 1, 'checkbox', N'["女性","男性","20歲↓","21~30歲","31~50歲","50歲↑"]', 3, 0),
(2, N'活動週期', 1, 1, 1, 'checkbox', N'["每月","每季","每半年","每年","節日","固定月日"]', 4, 0),
(2, N'活動類型', 1, 1, 1, 'checkbox', N'["免運","滿額贈","商品促銷","週年慶","抽獎活動"]', 5, 0),
(2, N'未來規劃及目標', 1, 1, 1, 'checkbox', N'["系統優化/改版","展店","新品上市","提升營業額%","海外出口","直播","社群經營","引進海外商品"]', 6, 0),

-- 金流狀況
(3, N'付款方式', 1, 1, 1, 'checkbox', N'["取貨付款佔%","線上刷卡佔%","代碼繳費佔%","第三方支付佔%","LINE PAY佔%","APPLE PAY佔%"]', 1, 1),

-- 物流狀況
(4, N'包裹宅配方式', 1, 1, 1, 'checkbox', N'["數網宅配佔%,元","黑貓佔,元","新竹貨運佔,元","大榮佔,元","宅配通佔,元","順豐佔%,元"]', 1, 1),
(4, N'透過哪家貨運將商品配送到物流中心？', 1, 1, 1, 'checkbox', N'["黑貓","新竹貨運","大榮","宅配通","順豐","信傳/籠","自有車隊"]', 2, 0),
(4, N'目前使用的超取通路有哪些？', 1, 1, 1, 'checkbox', N'["7-11","全家","萊爾富","OK","蝦皮店到店"]', 3, 0),
(4, N'未來會想增加的通路有哪些？', 1, 1, 1, 'checkbox', N'["蝦皮店到店","全家","萊爾富","OK","加油站","藥妝店"]', 4, 0),

-- 跨國佈局
(5, N'目前跨境經驗', 1, 1, 1, 'checkbox', N'["無","新馬佔","港澳佔","大陸佔","日本佔","韓國佔","泰國佔","菲律賓佔","越南佔","歐洲佔","美國佔","加拿大佔","紐澳佔"]', 1, 1),
(5, N'每月跨境訂單件數', 1, 1, 1, 'radio', N'["無","100件↓","101~500件","501~1000件","1001~2000件","2000件↑"]', 2, 0),
(5, N'經營跨境的困難以及希望得到什麼幫助？', 1, 1, 0, 'text', NULL, 3, 0),
(5, N'覺得國內及國外消費者的需求差異為何？', 1, 1, 0, 'text', NULL, 4, 0),
(5, N'未來跨境展望及規劃？', 1, 1, 0, 'text', NULL, 5, 0),

-- 廣告採買
(6, N'廣告操作人', 1, 1, 1, 'radio', N'["無","自操佔%","代操佔%"]', 1, 1),
(6, N'廣告操作方式', 1, 1, 1, 'checkbox', N'["社群(FB/IG)","口碑","圖像","影音","關鍵字"]', 2, 0),
(6, N'每月廣告預算', 1, 1, 1, 'radio', N'["無","10萬↓","11~100萬","101~200萬","201~300萬","300萬↑"]', 3, 0),
(6, N'有沒有想嘗試的廣告方式？', 1, 1, 1, 'checkbox', N'["無","社群(FB/IG)","口碑","圖像","影音","關鍵字"]', 4, 0),

-- 行銷配合
(7, N'是否有意願使用Shopmore APP？', 1, 1, 1, 'radio', N'["是","否"]', 1, 0),
(7, N'是否有意願合作預購檔期？', 1, 1, 1, 'radio', N'["是","否"]', 2, 0),
(7, N'是否有自己的Line OA？有什麼功能？', 1, 1, 1, 'checkbox', N'["優惠券","集點卡","問卷調查","主動推播","功能選單","分眾標籤"]', 3, 0),

-- 新服務
(8, N'數網提供新服務之合作意願？', 1, 1, 1, 'checkbox', N'["宅配經銷","上收服務","集運進口","出口","進口","進口商品批發","無意願"]', 1, 0),

-- 經驗分享
(9, N'覺得哪一家超取通路做得好？', 1, 1, 1, 'radio', N'["7-11","全家","萊爾富","OK","蝦皮店到店"]', 1, 0),
(9, N'特色優點', 1, 1, 1, 'checkbox', N'["價格便宜","異常件較少","行銷活動多","資訊能力強"]', 2, 0),
(9, N'最近同業電商狀況如何？', 1, 1, 0, 'text', NULL, 3, 0),
(9, N'對於電商明年整體市場看法及發展趨勢？', 1, 1, 1, 'radio', N'["大幅往上","微幅往上","持平","微幅往下","大幅往下"]', 4, 0),
(9, N'對於未來電商看好什麼類型的商品？', 1, 1, 0, 'text', NULL, 5, 0),
(9, N'經營平台是否有遇到困難？', 1, 1, 1, 'checkbox', N'["配送時效過長","異常件問題","整體市場下滑","配送費用過高","退貨率/未取率過高","供應商管理","進出口問題"]', 6, 0),
(9, N'對於實體門市未來發展規劃？', 1, 1, 1, 'radio', N'["無規劃","有"]', 7, 0)
GO

-- 建立訪廠答案歷史記錄表
CREATE TABLE [dbo].[VisitAnswerHistory](
    [HistoryID] [int] IDENTITY(1,1) PRIMARY KEY,
    [QuestionID] [int] NOT NULL,
    [VendorID] [int] NOT NULL,
    [Answer] [nvarchar](max) NOT NULL,
    [CreatedBy] [int] NOT NULL,
    [CreatedDate] [datetime] DEFAULT GETDATE(),
    FOREIGN KEY (QuestionID) REFERENCES VisitQuestions(QuestionID),
    FOREIGN KEY (VendorID) REFERENCES Vendors(VendorID),
    FOREIGN KEY (CreatedBy) REFERENCES Users(UserID)
)
GO

-- 建立訪廠問題資料表
CREATE TABLE Questions (
    QuestionID INT IDENTITY(1,1) PRIMARY KEY,
    QuestionText NVARCHAR(500) NOT NULL,
    Category NVARCHAR(50),
    SortOrder INT,
    IsActive BIT DEFAULT 1,
    CreatedDate DATETIME DEFAULT GETDATE(),
    ModifiedDate DATETIME DEFAULT GETDATE()
);

-- 插入預設問題
INSERT INTO Questions (QuestionText, Category, SortOrder) VALUES 
(N'公司營運狀況', N'基本資料', 1),
(N'主要產品與服務', N'基本資料', 2),
(N'品質管理系統認證', N'品質系統', 3),
(N'生產設備與產能', N'生產管理', 4),
(N'研發能力與未來規劃', N'研發創新', 5);
GO 