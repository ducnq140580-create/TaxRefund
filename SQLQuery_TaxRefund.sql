CREATE DATABASE TAXREFUND
GO

use TAXREFUND
GO

DROP TABLE #TempDestinationTable;

WITH DailyCounts AS (
    SELECT 
        CAST(NgayHT AS DATE) AS ExecutionDate,
        SoHC,
        COUNT(DISTINCT SoHC) AS DailyOccurrence
    FROM [dbo].[BC09]
    WHERE NgayHT BETWEEN '2025-10-01' AND '2025-12-02' 
      AND SoHC IS NOT NULL
    GROUP BY CAST(NgayHT AS DATE), SoHC
)
-- The INTO clause must come after the SELECT and before the FROM
SELECT 
    ExecutionDate,
    SoHC,
    DailyOccurrence,
    SUM(DailyOccurrence) OVER(PARTITION BY SoHC) AS TotalOccurrenceInPeriod
INTO #TempDestinationTable
FROM DailyCounts
ORDER BY SoHC, ExecutionDate;

-- Verify the result
SELECT * FROM #TempDestinationTable
ORDER BY TotalOccurrenceInPeriod Desc
GO ----

DROP TABLE #TempDestinationTable;

WITH DailyCounts AS (
    SELECT 
        CAST(NgayHT AS DATE) AS ExecutionDate,
        SoHC,
        COUNT(DISTINCT SoHC) AS DailyOccurrence
    FROM [dbo].[BC09]
    WHERE NgayHT BETWEEN '2025-10-01' AND '2025-12-02' 
      AND SoHC IS NOT NULL
    GROUP BY CAST(NgayHT AS DATE), SoHC
)
-- The INTO clause must come after the SELECT and before the FROM
SELECT    
    DISTINCT SoHC,    
    SUM(DailyOccurrence) OVER(PARTITION BY SoHC) AS TotalOccurrenceInPeriod
INTO #TempDestinationTable
FROM DailyCounts

-- Verify the result
SELECT * FROM #TempDestinationTable
ORDER BY TotalOccurrenceInPeriod Desc

SELECT b.SoHD, b.NgayHD, b.KyhieuHD, b.HoTenHK, b.SoHC, b.NgayHC, b.Quoctich,
       b.NgayHT, b.TrigiaHHchuaVAT, b.SotienVATDH, b.SotienDVNHH, 
       b.MasoDN, b.TenDNBH, b.Ghichu
FROM BC09 as b 
INNER JOIN #TempDestinationTable as tmpt ON b.SoHC = tmpt.SoHC
WHERE tmpt.TotalOccurrenceInPeriod >= 6 AND (b.NgayHT BETWEEN '2025-10-01' AND '2025-12-02')
ORDER BY b.SoHC;
GO---

Select *
From #TempDestinationTable
GO

Select count(SoHC)
From BC09
WHERE NgayHT BETWEEN N'2025-12-01' AND N'2025-12-02'

Select ROW_NUMBER() OVER (ORDER BY LoginName DESC) AS [STT], LoginName as [User], TenCC as [Officer Name], 
LoginPassword as [Password], Decentralization, NgaynhapHT 
From Login

Select LoginPassword 
From Login 
Where LoginName = N'HQ10-0152'
GO

Select distinct MasoDN
from BC09
order by MasoDN desc
GO

Select * 
From BC09
Where NgayHT = '2025-03-30'
GO

SELECT KyhieuHD, SoHD, NgayHD, TenDNBH, MasoDN, HoTenHK, SoHC, NgayHC, Quoctich, TrigiaHHchuaVAT, 
NgayHT, SotienVATDH, SotienDVNHH, (SELECT Distinct bc03.ThoigianGD 
FROM BC03 as bc03 INNER JOIN BC04 as bc04 ON bc03.MasoGD = bc04.MasoGD WHERE bc04.KyhieuSoNgay = '1C25TRF/75/18-02-2025') as NgaygioHT
FROM BC09 
WHERE KyhieuHD = '1C25TRF' AND SoHD = '75' AND NgayHD = N'2025-02-18' AND SoHC = 'M07018291' 
AND NgayHT = N'2025-03-30'
GO

Select distinct  ROW_NUMBER() OVER (ORDER BY bc09.NgayHD ASC) AS [STT], bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH                        
From BC09 as bc09
Where bc09.NgayHD between N'2025-04-30' and N'2025-05-08'

Select Top 10 *
From BC03
Order by NgaynhapHT desc
GO

Select Top 10 *
From BC04
Order by NgaynhapHT desc
GO


SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.NgayHD,  bc09.KyhieuHD,  bc09.HoTenHK,   bc09.SoHC,  bc09.NgayHC,  bc09.Quoctich,   
                                bc09.NgayHT,  bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, bc09.MasoDN,  bc09.TenDNBH --, bc03.ThoigianDD 
								FROM BC09 as bc09 LEFT JOIN BC04 as bc04 ON 
								bc09.SoHD = SUBSTRING(bc04.KyhieuSoNgay, 
								CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
								CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)
                            LEFT JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD
                            WHERE bc09.NgayHT between N'2025-07-28' and N'2025-07-29'
                            ORDER BY bc09.NgayHT DESC

SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.NgayHD,  bc09.KyhieuHD,  bc09.HoTenHK,   bc09.SoHC,  bc09.NgayHC,  bc09.Quoctich,   
                                bc09.NgayHT,  bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, bc09.MasoDN,  bc09.TenDNBH --, bc03.ThoigianDD 
								FROM BC09 as bc09 LEFT JOIN BC04 as bc04 ON 
								bc09.SoHD = SUBSTRING(bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay), 
								CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)
                            LEFT JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD
                            WHERE bc09.SoHD = '1204' AND (bc09.NgayHT between N'2025-07-28' and N'2025-07-29')
                            ORDER BY bc09.NgayHT DESC

SELECT SUBSTRING(bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1, LEN(bc04.KyhieuSoNgay) - 10)
as SoHC
FROM BC04 as bc04

DECLARE @str VARCHAR(100) = '1C25TAD/43/17-01-2025';

SELECT SUBSTRING(
    bc04.KyhieuSoNgay, 
    CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
    CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1
) AS Result
FROM BC04 as bc04


SELECT 
    CASE 
        WHEN bc04.KyhieuSoNgay LIKE '%/%/%' THEN
            SUBSTRING(
                bc04.KyhieuSoNgay, 
                CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
                CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) 
                - CHARINDEX('/', bc04.KyhieuSoNgay) - 1
            )
        ELSE NULL 
    END AS Result
FROM BC04 AS bc04;

SELECT CHARINDEX('1204', '1C25MRD/1204/21-07-2025'); 

SELECT SUBSTRING('1C25TAD/43/17-01-2025', CHARINDEX('/', '1C25TAD/43/17-01-2025') + 1, 
CHARINDEX('/', '1C25TAD/43/17-01-2025') + 1 - CHARINDEX('/', '1C25TAD/43/17-01-2025') + ) AS SoHD

SELECT SUBSTRING('1C25MRD/1204/21-07-2025', CHARINDEX('/', '1C25MRD/1204/21-07-2025') + 1, 
LEN('1C25MRD/1204/21-07-2025') - 19) AS SoHD

SELECT SUBSTRING('1C25TAD/43/17-01-2025', 8, 
LEN('1C25TAD/43/17-01-2025') - 19) AS SoHD

SELECT SUBSTRING(KyhieuSoNgay, CHARINDEX('/', KyhieuSoNgay ) + 1, 
LEN(KyhieuSoNgay) - 11) AS SoHD
FROM BC04 

SELECT SUBSTRING(KyhieuSoNgay, 9, 
LEN(KyhieuSoNgay) - 19) AS SoHD, KyhieuSoNgay, MasoGD
FROM BC04 
WHERE SUBSTRING(KyhieuSoNgay, 9, 
LEN(KyhieuSoNgay) - 19) = N'43' AND KyhieuSoNgay LIKE '%/%/%'

SELECT 
    SUBSTRING(KyhieuSoNgay, 9, LEN(KyhieuSoNgay) - 19) AS SoHD, 
    KyhieuSoNgay, 
    MasoGD
FROM BC04 
WHERE (SUBSTRING(KyhieuSoNgay, 9, LEN(KyhieuSoNgay) - 19)) = N'43'

WITH CTE_Extracted AS (
    SELECT 
        SUBSTRING(KyhieuSoNgay, 9, LEN(KyhieuSoNgay) - 19) AS SoHD,
        KyhieuSoNgay,
        MasoGD
    FROM BC04
)
SELECT SoHD, KyhieuSoNgay, MasoGD
FROM CTE_Extracted
WHERE SoHD = N'43'

SELECT SoHD, KyhieuSoNgay, MasoGD
FROM (
    SELECT 
        SUBSTRING(KyhieuSoNgay, 9, LEN(KyhieuSoNgay) - 19) AS SoHD,
        KyhieuSoNgay,
        MasoGD
    FROM BC04
) AS DerivedTable
WHERE SoHD = N'43'

SELECT DISTINCT
                                  ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
                                  bc09.SoHD, bc09.NgayHD, bc09.KyhieuHD,
                                  bc04.KyhieuSoNgay, 
                                  bc03.ThoigianGD,
                                  bc09.NgayHT,
                                  bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, 
                                  bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH,
                                  bc09.MasoDN, bc09.TenDNBH
                              FROM BC09 as bc09
                              INNER JOIN BC04 as bc04 ON bc09.SoHD = SUBSTRING(
                                  bc04.KyhieuSoNgay, 9, LEN(bc04.KyhieuSoNgay) - 19)                                                                  
                              INNER JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD  
                          WHERE SUBSTRING(
                                  bc04.KyhieuSoNgay, 9, LEN(bc04.KyhieuSoNgay) - 19) = N'43' AND bc04.KyhieuSoNgay LIKE '%/%/%' 

SELECT SUBSTRING('1C25MRD/1204/21-07-2025', 9, LEN('1C25MRD/1204/21-07-2025') - CHARINDEX('/', REVERSE('1C25MRD/1204/21-07-2025') + 1)

DECLARE @str VARCHAR(50) = '1C25MRD/1204/21-07-2025';
SELECT SUBSTRING(
    @str, 
    CHARINDEX('/', @str) + 1, 
    CHARINDEX('/', @str, CHARINDEX('/', @str) + 1) - CHARINDEX('/', @str) - 1
) AS SecondPart;

SELECT 
    SUBSTRING(
        KyhieuSoNgay, 
        CHARINDEX('/', KyhieuSoNgay) + 1, 
        CHARINDEX('/', KyhieuSoNgay, CHARINDEX('/', KyhieuSoNgay) + 1) - CHARINDEX('/', KyhieuSoNgay) - 1
    ) AS ExtractedValue --, NgayHT, MasoDD
FROM BC04
WHERE NgayHT between N'2025-07-28' and N'2025-07-28' --KyhieuSoNgay LIKE '%/1024/%' --AND (NgayHT between N'2025-07-28' and N'2025-07-30');
ORDER BY NgayHT desc
 
 WITH SplitValues AS (
    SELECT 
        value,
        ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS Position
    FROM BC04
    CROSS APPLY STRING_SPLIT(KyhieuSoNgay, '/')
    WHERE KyhieuSoNgay LIKE '%/%/%'
)
SELECT 
    MAX(CASE WHEN Position = 2 THEN value END) AS MiddleValue
FROM SplitValues
GROUP BY (SELECT 1); -- Dummy group by for aggregation

SELECT 
    PARSENAME(REPLACE(REPLACE(KyhieuSoNgay, '/', '.'), '-', '.'), 3) AS MiddleValue

SELECT 
    KyhieuSoNgay,
    SUBSTRING(
        KyhieuSoNgay,
        CHARINDEX('/', KyhieuSoNgay) + 1,
        CHARINDEX('/', KyhieuSoNgay, CHARINDEX('/', KyhieuSoNgay) + 1) - CHARINDEX('/', KyhieuSoNgay) - 1
    ) AS SecondPart
FROM [dbo].[BC04]
WHERE KyhieuSoNgay LIKE '%/%/%';  -- Only rows with at least two slashes

SELECT 
                                ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
                                bc09.SoHD, 
                                bc09.NgayHD, 
                                bc09.KyhieuHD, 
                                bc09.HoTenHK, 
                                bc09.SoHC, 
                                bc09.NgayHC, 
                                bc09.Quoctich, 
                                bc09.NgayHT,								
                                bc03.ThoigianGD,
                                bc09.TrigiaHHchuaVAT, 
                                bc09.SotienVATDH, 
                                bc09.SotienDVNHH,
                                bc09.MasoDN, 
                                bc09.TenDNBH
                            FROM BC09 as bc09
                            INNER JOIN BC04 as bc04 ON 
                                bc09.SoHD = SUBSTRING(
                                    bc04.KyhieuSoNgay, 
                                    CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
                                    CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)                                    
                            INNER JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD
                            WHERE bc09.SoHD = N'1204' AND KyhieuSoNgay LIKE '%/%/%';

FROM BC04
WHERE NgayHT between N'2025-07-28' and N'2025-07-29' AND KyhieuSoNgay LIKE '%/%/%';

Select Distinct NgaynhapHT 
From BC09 
Where NgayHT Between N'" + rfdate + "' And N'" + rtdate + "'

Insert Into tmpBC03 
Select ThoigianDD, MasoDD, HotenHK, SotienVNDHT, Ghichu, NgaynhapHT, LoginName 
From BC03 
Where (NgaynhapHT between N'2024-10-30' and N'2024-10-31') And LoginName = N'HQ10-0152'
GO

select *
from tmpBC09
where NgaynhapHT between N'2024-10-30' and N'2024-10-31'
Order by NgaynhapHT desc
GO

select *
from BC04
where NgaynhapHT between N'2024-10-30' and N'2024-10-31'
Order by NgaynhapHT desc
GO

select *
from BC04
where NgaynhapHT between N'2024-10-30' and N'2024-10-31'
Order by NgaynhapHT desc
GO

--Drop Table--

DROP TABLE [dbo].[Congchuc];
GO
----
DROP TABLE [dbo].[BC03];
GO
---
DROP TABLE [dbo].[tmpBC03];
GO
---
-- Create Table---

CREATE TABLE [dbo].[Congchuc](
	[SHCC] [varchar](9) NOT NULL,
	[TenCC] [nvarchar](50) NOT NULL,	
	[Ghichu] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[SHCC] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO
-------

CREATE TABLE [dbo].[tmpCongchuc](
	[SHCC] [varchar](9) NOT NULL,
	[TenCC] [nvarchar](50) NOT NULL,	
	[Ghichu] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[SHCC] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

Create table Login (
LoginName Varchar(9) NOT NULL,
TenCC Nvarchar (50),
LoginPassword Nvarchar(20) NOT NULL,
Decentralization bit,
NgaynhapHT Datetime,
Primary Key (LoginName)
)

GO

CREATE TABLE [dbo].[tmpBC03](
	[ThoigianDD] [datetime] NULL,
	[MasoDD] [nvarchar](6) NOT NULL,
	[HotenHK] [nvarchar](50) NOT NULL,
	[SotienVNDHT] [decimal](12) NULL,		
	[Ghichu] [nvarchar](max) NULL,
	[NgaynhapHT] [datetime] NULL,
	[LoginName] [varchar](9) NULL,
PRIMARY KEY CLUSTERED 
(
	[MasoDD] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

----
CREATE TABLE [dbo].[BC03](
	[ThoigianDD] [datetime] NULL,
	[MasoDD] [nvarchar](6) NOT NULL,
	[HotenHK] [nvarchar](50) NOT NULL,
	[SotienVNDHT] [decimal](12) NULL,		
	[Ghichu] [nvarchar](max) NULL,
	[NgaynhapHT] [datetime] NULL,
	[LoginName] [varchar](9) NULL,
PRIMARY KEY CLUSTERED 
(
	[MasoDD] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

----

CREATE TABLE [dbo].[tmpBC04](
	[KyhieuSoNgay] [varchar](50) NOT NULL,	
	[MasoDD][nvarchar](6) NULL,
	[TenDNBH][nvarchar](max) NULL,
	[SotienVATHD] [decimal](12) NULL,	
	[NgayHT] [datetime] NULL,	
	[Ghichu] [nvarchar](max) NULL,
	[NgaynhapHT] [datetime] NULL,
	[LoginName] [varchar](9) NULL,)
--PRIMARY KEY CLUSTERED 
--(
--	[KyhieuSoNgay] ASC
--)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
--) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

CREATE TABLE [dbo].[BC04](
	[KyhieuSoNgay] [varchar](50) NOT NULL,	
	[MasoDD][nvarchar](6) NULL,
	[TenDNBH][nvarchar](max) NULL,
	[SotienVATHD] [decimal](12) NULL,	
	[NgayHT] [datetime] NULL,	
	[Ghichu] [nvarchar](max) NULL,
	[NgaynhapHT] [datetime] NULL,
	[LoginName] [varchar](9) NULL,)
--PRIMARY KEY CLUSTERED 
--(
--	[KyhieuSoNgay] ASC
--)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
--) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

CREATE TABLE [dbo].[tmpBC09](
	--[Id] [int] IDENTITY(1,1) NOT NULL,
	[KyhieuHD] [varchar](8) NOT NULL,
	[SoHD] [varchar](8) NOT NULL,
	[NgayHD] [datetime] NULL,	
	[TenDNBH][nvarchar](max) NULL,	
	[MasoDN][varchar](14) NULL,
	[HoTenHK][nvarchar](max) NULL,
	[SoHC][varchar](12) NULL,
	[NgayHC] [datetime] NULL,
	[Quoctich] [varchar](3) NULL,
	[TrigiaHHchuaVAT] [decimal](12) NULL,	
	[NgayHT] [datetime] NULL,
	[SotienVATDH] [decimal](12) NULL,	
	[SotienDVNHH] [decimal](12) NULL,		
	[Ghichu] [nvarchar](max) NULL,
	[NgaynhapHT] [datetime] NULL,
	[LoginName] [varchar](9) NULL,)
--PRIMARY KEY CLUSTERED 
--(
--	--[Id] ASC
--)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
--) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

--GO

SET ANSI_PADDING OFF
GO

CREATE TABLE [dbo].[BC09](
	--[Id] [int] IDENTITY(1,1) NOT NULL,
	[KyhieuHD] [varchar](8) NOT NULL,
	[SoHD] [varchar](8) NOT NULL,
	[NgayHD] [datetime] NULL,	
	[TenDNBH][nvarchar](max) NULL,	
	[MasoDN][varchar](14) NULL,
	[HoTenHK][nvarchar](max) NULL,
	[SoHC][varchar](12) NULL,
	[NgayHC] [datetime] NULL,
	[Quoctich] [varchar](3) NULL,
	[TrigiaHHchuaVAT] [decimal](12) NULL,	
	[NgayHT] [datetime] NULL,
	[SotienVATDH] [decimal](12) NULL,	
	[SotienDVNHH] [decimal](12) NULL,		
	[Ghichu] [nvarchar](max) NULL,
	[NgaynhapHT] [datetime] NULL,
	[LoginName] [varchar](9) NULL,)
--PRIMARY KEY CLUSTERED 
--(
--	--[Id] ASC
--)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
--) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

--GO

SET ANSI_PADDING OFF
GO

select * 
from tmpBC03
go

select * 
from BC03
--Where MasoDD between '58006' and '59568'
order by MasoDD
go

select * 
from tmpBC04
--Where MasoDD between '58006' and '59568'
order by MasoDD
go

select * 
from BC04
--Where NgayHT between '2023-12-01' and '2023-12-31'
order by NgayHT
go

select * 
from tmpBC09
Where NgayHT between '2023-12-01' and '2023-12-31'
go

select * 
from BC09
Where NgayHT between '2023-12-01' and '2023-12-31'
go

------****************----
delete 
from BC03
Where ThoigianDD between '2023-12-01' and '2023-12-31'
go

delete
from BC04
go

delete
from BC09
Where NgayHT between '2023-12-01' and '2023-12-31'
go

----&&&&&&&&&&&&&&&&&&&&&&---------
SELECT * 
  FROM TestBC09

select *
from BC09
--Where NgayHT between '2023-12-01' and '2023-12-31'
order by NgayHT
GO

select *
from TestBC09 as t 
--Inner Join TestBC09 as t On b.SoHD = t.F3
Where t.F3 Not In (select distinct b.SoHD From BC09 as b Inner Join TestBC09 as t On b.SoHD = t.F3 and b.KyhieuHD = t.F2 and 
(b.NgayHT between '2023-12-01' and '2023-12-31'))
--order by b.NgayHT
GO


Select Distinct b.MasoDD, b.ThoigianDD From BC03 as b 
Inner Join tmpBC03 as tmpb 
On b.MasoDD = tmpb.MasoDD and b.ThoigianDD = tmpb.ThoigianDD
GO

Insert Into BC03 
Select Distinct tmpb.ThoigianDD, tmpb.MasoDD, tmpb.HotenHK, tmpb.SotienVNDHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName 
From tmpBC03 as tmpb 
Where tmpb.MasoDD Not In (Select Distinct b.MasoDD From BC03 as b Inner Join tmpBC03 as tmpb On b.MasoDD = tmpb.MasoDD and b.ThoigianDD = tmpb.ThoigianDD)
GO

Insert Into BC03 
Select Distinct tmpb.ThoigianDD, tmpb.MasoDD, tmpb.HotenHK, tmpb.SotienVNDHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName
From tmpBC03 as tmpb
GO

Select Distinct b.KyhieuSoNgay, b.MasoDD, b.TenDNBH, b.SotienVATHD, b.NgayHT, b.Ghichu, b.NgaynhapHT, b.LoginName 
From BC04 as b Inner Join tmpBC04 as tmpb On (b.KyhieuSoNgay = tmpb.KyhieuSoNgay)
Order by tmpb.MasoDD
GO

Select ROW_NUMBER() OVER (ORDER BY bc09.NgayHD DESC) AS [STT], bc09.SoHD, bc09.NgayHD, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH                        
From BC09 as bc09
Where bc09.SoHC = 'M63800413'
--Where utnv.LoginName = N'" + loginName + "'
GO

Select Top 20 ROW_NUMBER() OVER (ORDER BY bc09.NgayHD DESC) AS [STT], bc09.SoHD, bc09.NgayHD, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH                        
From BC09 as bc09
Where bc09.SotienVATDH > 100000000
GO

Select ROW_NUMBER() OVER (ORDER BY bc09.NgayHD ASC) AS [STT], bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH                        
From BC09 as bc09
Where bc09.NgayHT between '2023-11-01' and '2023-11-30'
GO

select SUM(bc09.TrigiaHHchuaVAT) as TongtrigiaHHchuaVAT
From BC09 as bc09
Where bc09.NgayHT between '2023-11-16' and '2023-11-27'
GO

SELECT distinct bc04.KyhieuSoNgay, bc03.ThoigianDD as NgayHT, bc03.MasoDD, bc03.HotenHK, bc03.SotienVNDHT
FROM BC03 as bc03 Inner Join BC04 as bc04 on bc03.MasoDD = bc04.MasoDD 
Where SotienVNDHT > '500000000'



Select ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.NgayHD, 
bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH 
From BC09 as bc09 Inner Join BC04 as bc04 
Where bc09.SotienVNDHT = '100000000'
GO


SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[DupplicateToNowSearch_sp]
	-- Add the parameters for the stored procedure here
	@tungay Datetime, @hientai Datetime
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;   

Select distinct  ROW_NUMBER() OVER (ORDER BY bc09.NgayHD ASC) AS [STT], bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH                        
From BC09 as bc09
Where (bc09.NgayHD between @tungay and @hientai) 
END

CREATE PROCEDURE [dbo].[DupplicateSearch_sp]
	-- Add the parameters for the stored procedure here
	@tungay Datetime, @denngay Datetime
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;   

Select distinct  ROW_NUMBER() OVER (ORDER BY bc09.NgayHD ASC) AS [STT], bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH                        
From BC09 as bc09
Where (bc09.NgayHD between @tungay and @denngay) 
END

==============================

CREATE PROCEDURE [dbo].[RefundManyTimesToNowSearch_sp]
	-- Add the parameters for the stored procedure here
	@rfdate Datetime, @hientai Datetime, @solanHT int
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;   

DROP TABLE [dbo].[#TempDestinationTable];
 

Select SoHC, COUNT(SoHC) as SolanHT
Into #TempDestinationTable
From BC09 as bc09
Where bc09.NgayHT  Between N'2025-12-01' And N'2026-01-09'
Group by SoHC 
Order by SolanHT desc

select b.KyhieuHD, b.SoHD, b.NgayHD, b.TenDNBH, b.MasoDN, b.HoTenHK, b.SoHC, b.NgayHC, b.Quoctich,
b.TrigiaHHchuaVAT, b.NgayHT, b.SotienVATDH, b.SotienDVNHH, b.Ghichu
From BC09 as b Inner Join #TempDestinationTable as tmpt On  b.SoHC =  tmpt.SoHC
Where tmpt.SolanHT > @solanHT and (b.NgayHT between @rfdate and @hientai) 
Order by SoHC
END

Select * 
From #TempDestinationTable
GO

----------------------------

CREATE PROCEDURE [dbo].[RefundManyTimesSearch_sp]
	-- Add the parameters for the stored procedure here
	@rfdate Datetime, @rtdate Datetime, @solanHT int
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON; 

Select SoHC, COUNT(SoHC) as SolanHT
Into #TempDestinationTable
From BC09 as bc09
Where bc09.NgayHT between @rfdate and @rtdate 
Group by SoHC 
Order by SolanHT desc

select b.KyhieuHD, b.SoHD, b.NgayHD, b.TenDNBH, b.MasoDN, b.HoTenHK, b.SoHC, b.NgayHC, b.Quoctich,
b.TrigiaHHchuaVAT, b.NgayHT, b.SotienVATDH, b.SotienDVNHH, b.Ghichu
From BC09 as b Inner Join #TempDestinationTable as tmpt On  b.SoHC =  tmpt.SoHC
Where tmpt.SolanHT > @solanHT and (b.NgayHT between @rfdate and @rtdate) 
Order by SoHC
END

--------
CREATE PROCEDURE [dbo].[ReportSearch_sp]
	-- Add the parameters for the stored procedure here
	 @previousrfdate Datetime, @rtdate Datetime
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON; 

select  bc09.SoHD, 
bc09.SoHC, bc09.NgayHT, bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH 
From BC09 as bc09 
Where bc09.NgayHT between @previousrfdate and @rtdate 
Order by bc09.NgayHT
END

---------------
CREATE PROCEDURE [dbo].[ImportUser_sp]
	-- Add the parameters for the stored procedure here
	@sohieucc varchar, @tencc nvarchar, @ghichu nvarchar
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON; 

Create Table #TempDestinationTable 
	([SHCC] [varchar](9) NOT NULL,
	[TenCC] [nvarchar](50) NOT NULL,	
	[Ghichu] [nvarchar](max) NULL)

Insert into #TempDestinationTable (SHCC, TenCC, Ghichu) values (@sohieucc, @tencc, @ghichu)

Insert Into Congchuc Select Distinct t.SHCC, t.TenCC, t.Ghichu
From #TempDestinationTable as t
Where t.SHCC not in (select distinct t.SHCC from #TempDestinationTable as t inner join Congchuc as c on t.SHCC = c.SHCC)

Drop Table  #TempDestinationTable 
END
----------------------------

select * 
from #TempDestinationTable 

delete from Congchuc
where SHCC = 'H' 


select sum(TrigiaHHchuaVAT) as TongTrigiaHHchuaVAT, sum(SotienVATDH) as TongSotienVATDH, count(SoHC) as TongSoluotHKHT
from BC09
where NgayHT between '2023-11-26' and '2023-11-27'



Select ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.NgayHD, 
bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH 
From BC09 as bc09 Inner Join BC04 as bc04 
Where bc09.SotienVNDHT = '100000000'
GO



Select SoHC, COUNT(SoHC) as SolanHT
Into #TempDestinationTable
From BC09 as bc09
Where bc09.NgayHT between '2023-01-01' and '2023-11-27'
Group by SoHC 
Order by SolanHT desc
GO

select  ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.NgayHD, 
bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT, bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH 
From BC09 as bc09 Inner Join #TempDestinationTable as tmpt On  bc09.SoHC =  tmpt.SoHC
Where tmpt.SolanHT > 15 and (bc09.NgayHT between '2023-01-01' and '2023-11-27') 
Order by SoHC
GO



SELECT  [Nhom]
      ,[Giovaoca]
      ,[Gioraca]
  FROM [ECUSAUDIT].[dbo].[LICHTRUC]

SELECT  [SOHD]
      ,[NGAYHT]
  FROM [ECUSAUDIT].[dbo].[NGAYHT]

select SOHD, Nhom, Giovaoca, Gioraca
from LICHTRUC as l inner join NGAYHT as nht on nht.NGAYHT = l.Gioraca
 




delete From UserTokhai
GO

select *
From UserDanhmucQLRR
where HSKhaibao = '2620999090'
GO

select * 
From BieuthueNK1

select * 
From UserDanhmucQLRR
Where LoginName = N'HQ10-0152' and (HSKhaibao ='' or HSKiemtra ='')
GO

delete 
From UserDanhmucQLRR
Where LoginName = N'HQ10-0152' and (HSKhaibao ='' or HSKiemtra ='')
GO


select * From BieuthueXK

delete 
From BieuthueNK
GO

Insert Into TokhaiNghivan (Trangthai, SoTK, NgayTK, MaLoaihinh, MaDN, TenDN, TenDoitac, SttHang, MasoHS, TSKhaibao,
MotaHanghoa, TenhangKhaibao, HSKiemtra, TSPhanloai, ThongtinRuiro, Soluong, DVT, ManuocXX, TennuocXX,
Dongia, MaNT, TrigiaNT, TrigiaTT,
SoTBKQPL, NgayTBKQPL, NgaynhapHT, LoginName)
Select ttknv.Trangthai, ttknv.SoTK, ttknv.NgayTK, ttknv.MaLoaihinh, ttknv.MaDN, ttknv.TenDN, ttknv.TenDoitac, ttknv.SttHang, ttknv.MasoHS, ttknv.TSKhaibao, 
ttknv.MotaHanghoa, ttknv.TenhangKhaibao, ttknv.HSKiemtra, ttknv.TSPhanloai, ttknv.ThongtinRuiro, ttknv.Soluong, ttknv.DVT, ttknv.ManuocXX, ttknv.TennuocXX,
ttknv.Dongia, ttknv.MaNT, ttknv.TrigiaNT, ttknv.TrigiaTT,
ttknv.SoTBKQPL, ttknv.NgayTBKQPL, ttknv.NgaynhapHT, ttknv.LoginName   
From tmpTokhaiNghivan as ttknv 
Where ttknv.LoginName = 'HQ10-0152' and ttknv.NgaynhapHT = '2022-02-20 12:18:22.347'
GO

Update TokhaiNghivan
Set Trangthai = '0'
 Where Trangthai = '1'
 GO

Insert Into TokhaiNghivan (Trangthai, SoTK, NgayTK, MaLoaihinh, MaDN, TenDN, TenDoitac, SttHang, MasoHS, TSKhaibao, 
MotaHanghoa, TenhangKhaibao, HSKiemtra, TSPhanloai, ThongtinRuiro, Soluong, DVT, ManuocXX, TennuocXX, Dongia, MaNT, TrigiaNT, TrigiaTT, SoTBKQPL, NgayTBKQPL, NgaynhapHT, LoginName)
Select Trangthai, SoTK, NgayTK, MaLoaihinh, MaDN, TenDN, TenDoitac, SttHang, MasoHS, TSKhaibao, 
MotaHanghoa, TenhangKhaibao, HSKiemtra, TSPhanloai, ThongtinRuiro, Soluong, DVT, ManuocXX, TennuocXX, Dongia, MaNT, TrigiaNT, TrigiaTT, SoTBKQPL, NgayTBKQPL, NgaynhapHT, LoginName
From TokhaiNghivan1  
GO

delete 
from TokhaiNghivan
GO

Select uk.Trangthai, uk.MasoHS, btn.TNKUD as TSKhaibao, uk.MasoPhanloai, btnk.TNKUD as TSPhanloai, 
uk.TenhangKhaibao, uk.MotaHanghoa, uk.SoTK, uk.NgayTK, uk.SoYeucau, uk.NgayYeucau, uk.SoTBKQPL, uk.NgayTBKQPL, 
uk.SoTBKQPT, uk.NgayTBKQPT From UserKQPTPL as uk Inner Join BieuthueNK as btn On uk.MasoHS = btn.MasoHS 
Inner Join BieuthueNK as btnk On uk.MasoPhanloai = btnk.MasoHS 
Where uk.LoginName = 'HQ10-0152' and uk.MasoHS = uk.MasoPhanloai
ORDER BY uk.MasoHS desc
GO

select * 
From UserKQPTPL
where MasoHS = MasoPhanloai and LoginName = 'HQ10-0152' 
GO

delete 
from UserKQPTPL
where MasoHS = MasoPhanloai and LoginName = 'HQ10-0152' 
GO

delete from UserKQPTPL
where MotaHanghoa in (Select distinct uk.MotaHanghoa From UserKQPTPL as uk 
Inner Join BieuthueNK as btn On uk.MasoHS = btn.MasoHS 
Inner Join BieuthueNK as btnk On uk.MasoPhanloai = btnk.MasoHS 
Where uk.LoginName = 'HQ10-0152' and uk.MasoHS = uk.MasoPhanloai)
GO


select *
from UserKQPTPL
where LoginName = 'HQ10-0152'
order by Id
GO

select *
from UserKQPTPL
where LoginName = 'HQ10-0152' and MasoPhanloai = ''
order by Id
GO

delete from UserKQPTPL
where LoginName = 'HQ10-0152' and MasoPhanloai = ''
GO

select * from tmpKQPTPL
Order by MasoHS
GO

select distinct Trangthai, SoTK, MaDN, NgayTK, SttHang, 
MasoHS, MasoHSDC, TenhangKhaibao, SoYeucau, NgayYeucau, SoPhieuchuyen, NgayPhieuchuyen, SoTBKQPT, NgayTBKQPT, 
SoTBGN, NgayTBGN, MasoTBGN, Chuong98, SoTBKQPL, NgayTBKQPL, MotaHanghoa, MasoPhanloai
from tmpKQPTPL
Where Trangthai = '1'
Order by MasoHS
GO

select * from UserKQPTPL
GO

Select *
From BieuthueNK
Order by TT asc
GO

Select *
From BieuthueNK
where TT between '4321' and '6234'
Order by TT asc
GO

Delete From BieuthueNK
where TT between '4321' and '6234'
GO
 
select *
from UserDanhmucQLRR
where LoginName = 'HQ10-0057'
GO

delete from DanhmucQLRR
GO

Insert into DanhmucQLRR (Trangthai, MotaHanghoa, HSKhaibao, HSKiemtra, ThongtinRuiro, VanbanThamchieu, NgayVB, Ghichu, NgaynhapHT, LoginName)                                     
select distinct Trangthai, MotaHanghoa, HSKhaibao, HSKiemtra, ThongtinRuiro, VanbanThamchieu, NgayVB, Ghichu, NgaynhapHT, LoginName                             
from UserDanhmucQLRR
where LoginName = 'HQ10-0057'
GO


Insert into UserDanhmucQLRR (Trangthai, XN, MotaHanghoa, HSKhaibao, HSKiemtra, ThongtinRuiro, VanbanThamchieu, NgayVB, Ghichu, NgaynhapHT, LoginName)                                     
select Trangthai, XN, MotaHanghoa, HSKhaibao, HSKiemtra, ThongtinRuiro, VanbanThamchieu, NgayVB, Ghichu, NgaynhapHT, LoginName                             
from UserDanhmucQLRR2
where LoginName = 'HQ10-0152'
GO

Update UserDanhmucQLRR 
Set XN = 'X'
GO

update DanhmucQLRR set LoginName = 'HQ10-0152'
GO

delete from UserKQPTPL
GO

select *
from tmpDanhmucQLRR
GO

select *
from Login
GO

select * 
from tmpTokhai
GO

select * 
from UserTokhai
GO


select *
from TokhaiNghivan
Where Trangthai = '1'
GO

delete from TokhaiNghivan
Where Trangthai = '1'
GO

delete from TokhaiNghivan
where TokhaiNghivan
GO

select *
from TokhaiLuu
GO

select * 
from Congchuc
GO

delete from UserTokhai
GO

Select * From UserDanhmucQLRR
GO

delete from UserDanhmucQLRR
GO


Insert Into UserTokhai 
Select Distinct SoTK, NgayTK, MaDN, TenDN, TenDoitac, MaLoaihinh, SttHang, MasoHS, MotaHanghoa, 
Soluong, DVT, Soluong2, DVT2, ManuocXX, TennuocXX, Dongia, MaNT, DVT1, TrigiaNT, TrigiaTT, NgaynhapHT, LoginName 
From tmpTokhai Where SoTK is not null
GO

Select r.Trangthai, ut.SoTK, ut.NgayTK, ut.MaLoaihinh, ut.MaDN,  ut.TenDN, ut.TenDoitac, ut.SttHang, ut.MasoHS, ut.MotaHanghoa, ut.Soluong,
ut.DVT, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT 
From UserTokhai as ut Inner Join DanhmucQLRR as r On ut.MasoHS = '68029310'
Where r.Trangthai = '1' and ut.LoginName = 'HQ10-0152' 
GO



---create database Table---
CREATE TABLE tmptDulieuV5
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	SoTK Varchar(12) NOT NULL,
	MaHQ varchar(4),
	MaLoaihinh varchar(3),	
	NgayHangDenDi Datetime,
	NgayTK Datetime,

	NgayHTKT Datetime, 
	NgayGPH Datetime, 
	NgayGPHV5 Datetime, 
	NgayBQ Datetime, 
	NgayCPXN Datetime,
		
	Trangthai bit,
	LuongTK varchar(1),
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),
	SoVandon varchar(30),
	TenCCDKV5 Nvarchar(50),
    TenCCKHV5 Nvarchar(50),
    SHCCDKV5 varchar(9),
    SHCCKHV5 varchar(9),
    SHCC varchar(9),
	PhanGhichu Nvarchar(MAX), 	
	NgaynhapHT Datetime,
	LoginName varchar(9),		
)
GO

CREATE TABLE DulieuV5
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	SoTK Varchar(12) NOT NULL,
	MaHQ varchar(4),
	MaLoaihinh varchar(3),	
	NgayHangDenDi Datetime,
	NgayTK Datetime,

	NgayHTKT Datetime, 
	NgayGPH Datetime, 
	NgayGPHV5 Datetime, 
	NgayBQ Datetime, 
	NgayCPXN Datetime,
		
	Trangthai bit,
	LuongTK varchar(1),
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),
	SoVandon varchar(30),
	TenCCDKV5 Nvarchar(50),
    TenCCKHV5 Nvarchar(50),
    SHCCDKV5 varchar(9),
    SHCCKHV5 varchar(9),
    SHCC varchar(9),
	PhanGhichu Nvarchar(MAX), 	
	NgaynhapHT Datetime,
	LoginName varchar(9),		
)
GO

CREATE TABLE tmpTokhai
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	SoTK Varchar(12) NOT NULL,
	NgayTK Datetime,
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),	
	TenDoitac Nvarchar(MAX),
	MaLoaihinh varchar(3),
	SttHang varchar(10),
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX),
	Soluong varchar(12),
	DVT varchar(5),
	Soluong2 varchar(12),
	DVT2 varchar(2),
	ManuocXX varchar(2),
	TennuocXX varchar(10),
	Dongia varchar(12),
	MaNT varchar(3),
	DVT1 varchar(2),
	TrigiaNT varchar(13),
	TrigiaTT varchar(13),
	NgaynhapHT Datetime,
	LoginName varchar(9),		
)
GO

CREATE TABLE UserTokhai
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	SoTK Varchar(12) NOT NULL,
	NgayTK Datetime,
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),	
	TenDoitac Nvarchar(MAX),
	MaLoaihinh varchar(3),
	SttHang varchar(10),
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX),
	Soluong varchar(12),
	DVT varchar(5),
	Soluong2 varchar(12),
	DVT2 varchar(2),
	ManuocXX varchar(2),
	TennuocXX varchar(10),
	Dongia varchar(12),
	MaNT varchar(3),
	DVT1 varchar(2),
	TrigiaNT varchar(13),
	TrigiaTT varchar(13),
	NgaynhapHT Datetime,
	LoginName varchar(9),		
)
GO

use ECUSAUDIT
GO

CREATE TABLE TokhaiNghivan
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	Trangthai bit,
	SoTK Varchar(12) NOT NULL,
	NgayTK Datetime,
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),
	SoVandon varchar(30),	
	TenDoitac Nvarchar(MAX),
	MaLoaihinh varchar(3),
	SttHang varchar(10),
	MasoHS varchar(10),
	TSKhaibao varchar(25),
	MotaHanghoa Nvarchar(MAX),
	TenhangKhaibao Nvarchar(MAX),
	HSKiemtra varchar(10),
	TSPhanloai varchar(25),
	ThongtinRuiro Nvarchar(MAX),
	Soluong varchar(12),
	DVT varchar(5),
	ManuocXX varchar(2),		
	TennuocXX varchar(10),
	Dongia varchar(12),
	MaNT varchar(3),	
	TrigiaNT varchar(13),
	TrigiaTT varchar(13),
	SoTBKQPL varchar(12), 
	NgayTBKQPL Datetime,	
	PhanGhichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9)
)
GO

CREATE TABLE TokhaiNghivan1
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	Trangthai bit,
	SoTK Varchar(12) NOT NULL,
	NgayTK Datetime,
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),
	SoVandon varchar(30),		
	TenDoitac Nvarchar(MAX),
	MaLoaihinh varchar(3),
	SttHang varchar(10),
	MasoHS varchar(10),
	TSKhaibao varchar(25),
	MotaHanghoa Nvarchar(MAX),
	TenhangKhaibao Nvarchar(MAX),
	HSKiemtra varchar(10),
	TSPhanloai varchar(25),
	ThongtinRuiro Nvarchar(MAX),
	Soluong varchar(12),
	DVT varchar(5),
	ManuocXX varchar(2),		
	TennuocXX varchar(10),
	Dongia varchar(12),
	MaNT varchar(3),	
	TrigiaNT varchar(13),
	TrigiaTT varchar(13),
	SoTBKQPL varchar(12), 
	NgayTBKQPL Datetime,
	PhanGhichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9)
)
GO

CREATE TABLE tmpTokhaiNghivan
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	Trangthai bit,
	SoTK varchar(12) NOT NULL,
	NgayTK Datetime,
	MaDN Nvarchar(14),
	TenDN Nvarchar(MAX),
	SoVandon varchar(30),		
	TenDoitac Nvarchar(MAX),
	MaLoaihinh varchar(3),
	SttHang varchar(10),
	MasoHS varchar(10),
	TSKhaibao varchar(25),
	MotaHanghoa Nvarchar(MAX),
	TenhangKhaibao Nvarchar(MAX),
	HSKiemtra varchar(10),
	TSPhanloai varchar(25),
	ThongtinRuiro Nvarchar(MAX),
	Soluong varchar(12),
	DVT varchar(5),
	ManuocXX varchar(2),		
	TennuocXX varchar(10),
	Dongia varchar(12),
	MaNT varchar(3),	
	TrigiaNT varchar(13),
	TrigiaTT varchar(13),
	SoTBKQPL varchar(12), 
	NgayTBKQPL Datetime,
	PhanGhichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9)
)
GO

CREATE TABLE tmpDanhmucQLRR
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Trangthai bit,
	MotaHanghoa Nvarchar(MAX) NOT NULL,
	HSKhaibao varchar(10) NOT NULL,
	HSKiemtra varchar(10) NOT NULL,
	ThongtinRuiro Nvarchar(MAX) NOT NULL,
	VanbanThamchieu Nvarchar(250),
	NgayVB Datetime,
	Ghichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE DanhmucQLRR
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Trangthai bit,
	XN varchar(1) NOT NULL,
	MotaHanghoa Nvarchar(MAX) NOT NULL,
	HSKhaibao varchar(10) NOT NULL,
	HSKiemtra varchar(10) NOT NULL,
	ThongtinRuiro Nvarchar(MAX) NOT NULL,
	VanbanThamchieu Nvarchar(250),
	NgayVB Datetime,
	Ghichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE tmpDanhmucQLRR
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Trangthai bit,
	XN varchar(1) NOT NULL,
	MotaHanghoa Nvarchar(MAX) NOT NULL,
	HSKhaibao varchar(10) NOT NULL,
	HSKiemtra varchar(10) NOT NULL,
	ThongtinRuiro Nvarchar(MAX) NOT NULL,
	VanbanThamchieu Nvarchar(250),
	NgayVB Datetime,
	Ghichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE UserDanhmucQLRR
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Trangthai bit,
	XN varchar(1) NOT NULL,
	MotaHanghoa Nvarchar(MAX) NOT NULL,
	HSKhaibao varchar(10) NOT NULL,
	HSKiemtra varchar(10) NOT NULL,
	ThongtinRuiro Nvarchar(MAX) NOT NULL,
	VanbanThamchieu Nvarchar(250),
	NgayVB Datetime,
	Ghichu Nvarchar(MAX),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE DanhmucPLHH
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,

)
GO



CREATE TABLE tmpKQPTPL
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Trangthai bit,
	SoTK Varchar(12) NOT NULL,	
	MaDN Nvarchar(14),
	NgayTK Datetime,
	SttHang varchar(2),
	MasoHS varchar(10),
	MasoHSDC varchar(10),
	TenhangKhaibao Nvarchar(MAX),
	SoYeucau varchar(14),
	NgayYeucau Datetime,
	SoPhieuchuyen varchar (14),
	NgayPhieuchuyen Datetime,
	SoTBKQPT varchar (12),
	NgayTBKQPT Datetime,
	SoTBGN varchar (14),
	NgayTBGN Datetime,
	MasoTBGN varchar(10),
	Chuong98 varchar(10),
	SoTBKQPL varchar (12),
	NgayTBKQPL Datetime,
	MotaHanghoa Nvarchar(MAX),
	MasoPhanloai varchar(10),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE UserKQPTPL
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Trangthai bit,
	SoTK Varchar(12) NOT NULL,	
	MaDN Nvarchar(14),
	NgayTK Datetime,
	SttHang varchar(2),
	MasoHS varchar(10),
	MasoHSDC varchar(10),
	TenhangKhaibao Nvarchar(MAX),
	SoYeucau varchar(14),
	NgayYeucau Datetime,
	SoPhieuchuyen varchar (14),
	NgayPhieuchuyen Datetime,
	SoTBKQPT varchar (12),
	NgayTBKQPT Datetime,
	SoTBGN varchar (14),
	NgayTBGN Datetime,
	MasoTBGN varchar(10),
	Chuong98 varchar(10),
	SoTBKQPL varchar (12),
	NgayTBKQPL Datetime,
	MotaHanghoa Nvarchar(MAX),
	MasoPhanloai varchar(10),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE tmpBieuthueNK1
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	TT Int, 
	Phannhom varchar(5),
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX) NOT NULL,

	GoodsDescription Nvarchar(MAX) NOT NULL,
	DVT nvarchar(15),
	TNKTT varchar(5),
	TNKUD varchar(5),

	VAT varchar(5),	
	ACFTA varchar(25),
	ATIGA varchar(25),
	AJCEP varchar(25),

	VJEPA varchar(25),
	AKFTA varchar(25),	
	AANZFTA varchar(25),
	AIFTA varchar(25),

	VKFTA varchar(25),
	VCFTA varchar(25),
	VN_EAEU varchar(25),
	CPTPP varchar(25),

	AHKFTA varchar(25),
	EVFTA varchar(25),	
	Ghichu nvarchar(250),
	CSMathang nvarchar(Max),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE BieuthueNK
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	TT Int, 
	Phannhom varchar(5),
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX) NOT NULL,

	GoodsDescription Nvarchar(MAX) NOT NULL,
	DVT nvarchar(15),
	TNKTT varchar(25),
	TNKUD varchar(25),

	VAT varchar(5),	
	ACFTA varchar(25),
	ATIGA varchar(25),
	AJCEP varchar(25),

	VJEPA varchar(25),
	AKFTA varchar(25),	
	AANZFTA varchar(25),
	AIFTA varchar(25),

	VKFTA varchar(25),
	VCFTA varchar(25),
	VN_EAEU varchar(25),
	CPTPP varchar(25),

	AHKFTA varchar(25),
	EVFTA varchar(25),	
	Ghichu nvarchar(250),
	CSMathang nvarchar(Max),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE BieuthueNK
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	TT Int, 
	Phannhom varchar(5),
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX) NOT NULL,

	GoodsDescription Nvarchar(MAX) NOT NULL,
	DVT nvarchar(15),
	TNKTT varchar(25),
	TNKUD varchar(25),	
	CHUONGCT varchar(25), 
	 
	VAT varchar(5),	
	ACFTA varchar(25), 	
	ATIGA varchar(25),	
	AJCEP varchar(25),	 
	
	VJEPA varchar(25),
	AKFTA varchar(25),	
	AANZFTA varchar(25),
	AIFTA varchar(25),

	VKFTA varchar(25),
	VCFTA varchar(25),
	VN_EAEU varchar(25),
	CPTPP_IV varchar(25),
	CPTPP_V varchar(25),

	AHKFTA varchar(25),
	EVFTA varchar(25),
	VNCU varchar(25),
	CSMathang nvarchar(Max),	
	Ghichu nvarchar(250),	
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE tmpBieuthueXK
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	TT Int,	
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX) NOT NULL,
	GoodsDescription Nvarchar(MAX),
	DVT nvarchar(15),

	TSXKUD varchar(25),
	TSXK125 varchar(25),
	TSXK7_2020 varchar(25),	
	TSXK2022 varchar(25),
	TSXK7_2022 varchar(25),

	EVFTA_2020 varchar(25),
	EVFTA_2021 varchar(25),
	EVFTA_2022 varchar(25),
	UKVFTA_2021 varchar(25),	
	UKVFTA_2022 varchar(25),
		
	CPTPP_I varchar(25),
	CPTPP_II varchar(25),
	CPTPP_III varchar(25),
	CPTPP_IV varchar(25),
	CPTPP_V varchar(25),	

	Ghichu nvarchar(250),
	CSMathang nvarchar(Max),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

CREATE TABLE BieuthueXK
(
	[Id] [int] IDENTITY(1,1) PRIMARY KEY CLUSTERED,	
	TT Int,	
	MasoHS varchar(10),
	MotaHanghoa Nvarchar(MAX) NOT NULL,
	GoodsDescription Nvarchar(MAX),
	DVT nvarchar(15),

	TSXKUD varchar(25),
	TSXK125 varchar(25),
	TSXK7_2020 varchar(25),	
	TSXK2022 varchar(25),
	TSXK7_2022 varchar(25),

	EVFTA_2020 varchar(25),
	EVFTA_2021 varchar(25),
	EVFTA_2022 varchar(25),
	UKVFTA_2021 varchar(25),	
	UKVFTA_2022 varchar(25),
		
	CPTPP_I varchar(25),
	CPTPP_II varchar(25),
	CPTPP_III varchar(25),
	CPTPP_IV varchar(25),
	CPTPP_V varchar(25),		

	Ghichu nvarchar(250),
	CSMathang nvarchar(Max),
	NgaynhapHT Datetime,
	LoginName varchar(9),
)
GO

select *
from tmpBieuthueNK
GO

delete
from BieuthueXK
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[FilterTokhaiNghivan_sp ]
	-- Add the parameters for the stored procedure here
	@loginName varchar
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here	
	Select r.Trangthai, ut.SoTK, ut.NgayTK, ut.MaLoaihinh, ut.MaDN,  ut.TenDN, ut.TenDoitac, ut.SttHang, ut.MasoHS, ut.MotaHanghoa, ut.Soluong,
	ut.DVT, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT
	From UserTokhai as ut Inner Join DanhmucQLRR as r On ut.MasoHS = r.HSKhaibao and r.Trangthai = '1'
	--Where ut.LoginName = @loginName and r.Trangthai = '1'
END
GO

Select Distinct Top 200 ROW_NUMBER() OVER (Order by ur.Id asc) as STT, ur.Trangthai, ur.XN, ur.MotaHanghoa, ur.HSKhaibao, 
IIF(ur.XN = 'N', btnk.TNKUD, btxk.TSXKUD) as TSKhaibao, ur.HSKiemtra, 
IIF(ur.XN = 'N', btn.TNKUD, btx.TSXKUD) as TSKiemtra, ur.ThongtinRuiro, ur.VanbanThamchieu, ur.NgayVB, ur.Ghichu, ur.NgaynhapHT, ur.LoginName 
From UserDanhmucQLRR as ur 
Inner Join BieuthueNK as btnk On ur.HSKhaibao = btnk.MasoHS 
Inner Join BieuthueNK as btn On ur.HSKiemtra = btn.MasoHS
Inner Join BieuthueXK as btxk On ur.HSKhaibao = btxk.MasoHS
Inner Join BieuthueXK as btx On ur.HSKiemtra = btx.MasoHS  
Where ur.LoginName = 'HQ10-0152'
GO

Select Distinct Top 200 ROW_NUMBER() OVER (Order by ur.Id asc) as STT, ur.Trangthai, ur.XN, ur.MotaHanghoa, ur.HSKhaibao, 
 btxk.TSXKUD as TSKhaibao, ur.HSKiemtra, 
 btx.TSXKUD as TSKiemtra, ur.ThongtinRuiro, ur.VanbanThamchieu, ur.NgayVB, ur.Ghichu, ur.NgaynhapHT, ur.LoginName 
From UserDanhmucQLRR as ur 
Inner Join BieuthueXK as btxk On ur.HSKhaibao = btxk.MasoHS
Inner Join BieuthueXK as btx On ur.HSKiemtra = btx.MasoHS  
Where ur.LoginName = 'HQ10-0152'
GO

USE [ECUSAUDIT]
GO
/****** Object:  StoredProcedure [dbo].[AnalyseKQPTPL_sp]    Script Date: 28/11/2021 7:54:27 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
ALTER PROCEDURE [dbo].[AnalyseKQPTPLNK_sp] 
	-- Add the parameters for the stored procedure here
	@utloginName varchar (9), @ukloginName varchar (9)

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;	

    -- Insert statements for procedure here
	Select Distinct  ROW_NUMBER() OVER (ORDER BY ut.MasoHS ASC) AS [STT], uk.Trangthai, ut.SoTK, v5.MaHQ, ut.NgayTK, ut.MaLoaihinh, 
	v5.LuongTK, v5.TenCCDKV5, v5.TenCCKHV5, v5.NgayHTKT, v5.NgayGPH, v5.NgayGPHV5, v5.NgayBQ, v5.NgayCPXN, v5.PhanGhichu,
	ut.SttHang, ut.MasoHS, 

	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btnk.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btnk.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btnk.VJEPA,	
	IIF(ut.ManuocXX = 'KR', btnk.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btnk.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btnk.AHKFTA,  
	IIF(ut.ManuocXX = 'FR' OR ut.ManuocXX = 'DE' OR ut.ManuocXX = 'IT' OR ut.ManuocXX = 'BE' OR ut.ManuocXX = 'NL' OR ut.ManuocXX = 'LU' 
	OR ut.ManuocXX = 'IE' OR ut.ManuocXX = 'DK' OR ut.ManuocXX = 'GR' OR ut.ManuocXX = 'EG' OR ut.ManuocXX = 'ES' OR ut.ManuocXX = 'PT'	
	OR ut.ManuocXX = 'AT' OR ut.ManuocXX = 'SE'OR ut.ManuocXX = 'FI' OR ut.ManuocXX = 'CZ' OR ut.ManuocXX = 'HU' OR ut.ManuocXX = 'PL' 
	OR ut.ManuocXX = 'SK' OR ut.ManuocXX = 'SI' OR ut.ManuocXX = 'LT' OR ut.ManuocXX = 'LV' OR ut.ManuocXX = 'EE' OR ut.ManuocXX = 'MT' 
	OR ut.ManuocXX = 'CY' OR ut.ManuocXX = 'BG' OR ut.ManuocXX = 'RO' OR ut.ManuocXX = 'HR', btnk.EVFTA, btnk.TNKUD))))))) as TSKhaibao, 

	uk.MasoHS as HSKhaibao, uk.MasoPhanloai as HSKiemtra, 

	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btn.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btn.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btn.VJEPA,	
	IIF(ut.ManuocXX = 'KR', btn.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btn.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btn.AHKFTA, 

	IIF(ut.ManuocXX = 'FR' OR ut.ManuocXX = 'DE' OR ut.ManuocXX = 'IT' OR ut.ManuocXX = 'BE' OR ut.ManuocXX = 'NL' OR ut.ManuocXX = 'LU' 
	OR ut.ManuocXX = 'IE' OR ut.ManuocXX = 'DK' OR ut.ManuocXX = 'GR' OR ut.ManuocXX = 'EG' OR ut.ManuocXX = 'ES' OR ut.ManuocXX = 'PT'	
	OR ut.ManuocXX = 'AT' OR ut.ManuocXX = 'SE'OR ut.ManuocXX = 'FI' OR ut.ManuocXX = 'CZ' OR ut.ManuocXX = 'HU' OR ut.ManuocXX = 'PL' 
	OR ut.ManuocXX = 'SK' OR ut.ManuocXX = 'SI' OR ut.ManuocXX = 'LT' OR ut.ManuocXX = 'LV' OR ut.ManuocXX = 'EE' OR ut.ManuocXX = 'MT' 
	OR ut.ManuocXX = 'CY' OR ut.ManuocXX = 'BG' OR ut.ManuocXX = 'RO' OR ut.ManuocXX = 'HR', btn.EVFTA, btn.TNKUD))))))) as TSPhanloai, 

	ut.MotaHanghoa, uk.TenhangKhaibao, uk.MotaHanghoa as ThongtinRuiro, ut.MaDN, ut.TenDN, ut.TenDoitac, 
    ut.Soluong, ut.DVT, ut.ManuocXX, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT, 
	uk.SoTBKQPL, uk.NgayTBKQPL, ut.NgaynhapHT, ut.LoginName
	 
    From UserTokhai as ut Inner Join UserKQPTPL as uk On ut.MasoHS = uk.MasoHS Inner Join BieuthueNK as btnk On uk.MasoHS = btnk.MasoHS 
    Inner Join BieuthueNK as btn On uk.MasoPhanloai = btn.MasoHS
	Inner Join DulieuV5 as v5 On ut.SoTK = v5.SoTK 
	Where (ut.LoginName = @utloginName and uk.LoginName = @ukloginName) and uk.Trangthai = '1' and Left(ut.SoTK,1) = '1' --Order By ut.MasoHS ASC

END

Select Distinct uk.Trangthai, ut.SoTK, ut.NgayTK, ut.MaLoaihinh, ut.SttHang, ut.MasoHS, 
	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btnk.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btnk.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btnk.VJEPA,	IIF(ut.ManuocXX = 'KR', btnk.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btnk.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btnk.AHKFTA, btnk.TNKUD)))))) as TSKhaibao, 

	uk.MasoHS as HSKhaibao, uk.MasoPhanloai as HSKiemtra, 

	IIF(ut.ManuocXX = 'CN'OR ut.TennuocXX = 'CN', btn.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btn.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btn.VJEPA,	IIF(ut.ManuocXX = 'KR', btn.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btn.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btn.AHKFTA, btn.TNKUD)))))) as TSPhanloai, 

	ut.MotaHanghoa, uk.TenhangKhaibao, uk.MotaHanghoa as ThongtinRuiro, ut.MaDN, ut.TenDN, ut.TenDoitac, 
    ut.Soluong, ut.DVT, ut.ManuocXX, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT, 
	uk.SoTBKQPL, uk.NgayTBKQPL, ut.NgaynhapHT, ut.LoginName 
    From UserTokhai as ut Inner Join UserKQPTPL as uk On ut.MasoHS = uk.MasoHS 
	Inner Join BieuthueNK as btnk On ut.MasoHS = btnk.MasoHS 
    Inner Join BieuthueNK as btn On uk.MasoPhanloai = btn.MasoHS 
	Where (ut.LoginName = @utloginName and uk.LoginName = @ukloginName) and uk.Trangthai = '1' Order By ut.MasoHS ASC
	END
	GO

USE [ECUSAUDIT]
GO
/****** Object:  StoredProcedure [dbo].[AnalyseKQPTPL_sp]    Script Date: 28/11/2021 7:54:27 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
	CREATE PROCEDURE [dbo].[AnalyseCS_sp] 
	-- Add the parameters for the stored procedure here
	@utloginName varchar (9), @tnloginName varchar (9)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	Select Distinct ROW_NUMBER() OVER (ORDER BY ut.MasoHS ASC) AS [STT], tn.Trangthai, ut.SoTK,  v5.MaHQ, ut.NgayTK, ut.MaLoaihinh, 
	v5.LuongTK, v5.TenCCDKV5, v5.TenCCKHV5, v5.NgayHTKT, v5.NgayGPH, v5.NgayGPHV5, v5.NgayBQ, v5.NgayCPXN, v5.PhanGhichu,
	ut.SttHang, ut.MasoHS, 
	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btnk.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btnk.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btnk.VJEPA,	
	IIF(ut.ManuocXX = 'KR', btnk.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btnk.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btnk.AHKFTA,  
	IIF(ut.ManuocXX = 'FR' OR ut.ManuocXX = 'DE' OR ut.ManuocXX = 'IT' OR ut.ManuocXX = 'BE' OR ut.ManuocXX = 'NL' OR ut.ManuocXX = 'LU' 
	OR ut.ManuocXX = 'IE' OR ut.ManuocXX = 'DK' OR ut.ManuocXX = 'GR' OR ut.ManuocXX = 'EG' OR ut.ManuocXX = 'ES' OR ut.ManuocXX = 'PT'	
	OR ut.ManuocXX = 'AT' OR ut.ManuocXX = 'SE'OR ut.ManuocXX = 'FI' OR ut.ManuocXX = 'CZ' OR ut.ManuocXX = 'HU' OR ut.ManuocXX = 'PL' 
	OR ut.ManuocXX = 'SK' OR ut.ManuocXX = 'SI' OR ut.ManuocXX = 'LT' OR ut.ManuocXX = 'LV' OR ut.ManuocXX = 'EE' OR ut.ManuocXX = 'MT' 
	OR ut.ManuocXX = 'CY' OR ut.ManuocXX = 'BG' OR ut.ManuocXX = 'RO' OR ut.ManuocXX = 'HR', btnk.EVFTA, btnk.TNKUD))))))) as TSKhaibao, 

	tn.MasoHS as HSKhaibao, tn.HSKiemtra, 

	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btn.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btn.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btn.VJEPA,	
	IIF(ut.ManuocXX = 'KR', btn.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btn.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btn.AHKFTA, 

	IIF(ut.ManuocXX = 'FR' OR ut.ManuocXX = 'DE' OR ut.ManuocXX = 'IT' OR ut.ManuocXX = 'BE' OR ut.ManuocXX = 'NL' OR ut.ManuocXX = 'LU' 
	OR ut.ManuocXX = 'IE' OR ut.ManuocXX = 'DK' OR ut.ManuocXX = 'GR' OR ut.ManuocXX = 'EG' OR ut.ManuocXX = 'ES' OR ut.ManuocXX = 'PT'	
	OR ut.ManuocXX = 'AT' OR ut.ManuocXX = 'SE'OR ut.ManuocXX = 'FI' OR ut.ManuocXX = 'CZ' OR ut.ManuocXX = 'HU' OR ut.ManuocXX = 'PL' 
	OR ut.ManuocXX = 'SK' OR ut.ManuocXX = 'SI' OR ut.ManuocXX = 'LT' OR ut.ManuocXX = 'LV' OR ut.ManuocXX = 'EE' OR ut.ManuocXX = 'MT' 
	OR ut.ManuocXX = 'CY' OR ut.ManuocXX = 'BG' OR ut.ManuocXX = 'RO' OR ut.ManuocXX = 'HR', btn.EVFTA, btn.TNKUD))))))) as TSPhanloai, 

	ut.MotaHanghoa, tn.TenhangKhaibao, tn.ThongtinRuiro, ut.MaDN, ut.TenDN, ut.TenDoitac, 
    ut.Soluong, ut.DVT, ut.ManuocXX, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT, 
	tn.SoTBKQPL, tn.NgayTBKQPL, ut.NgaynhapHT, ut.LoginName 
    From UserTokhai as ut 
	Inner Join TokhaiNghivan as tn On ut.MasoHS = tn.MasoHS 
	Inner Join BieuthueNK as btnk On tn.MasoHS = btnk.MasoHS 
    Inner Join BieuthueNK as btn On tn.HSKiemtra = btn.MasoHS 
	Inner Join DulieuV5 as v5 On ut.SoTK = v5.SoTK 
	Where (ut.LoginName = @utloginName and tn.LoginName = @tnloginName) --and Left(ut.SoTK,1) = '1' --Order By ut.MasoHS ASC
END
GO

-- ================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[LoadSuspect_sp] 
	-- Add the parameters for the stored procedure here
	@tnloginName varchar (9)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	Select Distinct tn.Trangthai, tn.SoTK, tn.NgayTK, tn.MaLoaihinh, tn.MaDN, tn.TenDN, tn.TenDoitac, tn.SttHang, 
    tn.MasoHS, tn.TSKhaibao, tn.MotaHanghoa, tn.TenhangKhaibao, tn.HSKiemtra, tn.TSPhanloai, tn.Thongtinruiro, 
    tn.Soluong, tn.DVT, tn.ManuocXX, tn.TennuocXX, tn.Dongia, tn.MaNT, tn.TrigiaNT, tn.TrigiaTT, tn.NgaynhapHT 
    From TokhaiNghivan as tn 
	Where (tn.LoginName = @tnloginName) Order by tn.MasoHS Asc
END
GO

ALTER PROCEDURE [dbo].[AnalyseRRXK_sp]  
	-- Add the parameters for the stored procedure here
	@utloginName varchar (9), @urloginName varchar (9)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	-- Insert statements for procedure here
	Select Distinct  ROW_NUMBER() OVER (ORDER BY ut.MasoHS ASC) AS [STT], ur.Trangthai, ut.SoTK, v5.MaHQ, ut.NgayTK, ut.MaLoaihinh, 
	v5.LuongTK, v5.TenCCDKV5, v5.TenCCKHV5, v5.NgayHTKT, v5.NgayGPH, v5.NgayGPHV5, v5.NgayBQ, v5.NgayCPXN, v5.PhanGhichu,
	ut.SttHang, ut.MasoHS, 

	--IIF(ut.NgayTK between '2018-01-01' and '2020-07-09', btxk.TSXK125,
	--IIF( ut.NgayTK between '2020-07-10' and '2021-12-29', btxk.TSXK7_2020,  
	--IIF(ut.NgayTK between '2021-12-30' and '2022-06-29', btxk.TSXK2022, btxk.TSXK7_2022))) as TSKhaibao,

	IIF(ut.NgayTK < '2020-07-10', btxk.TSXK125,
	IIF(ut.NgayTK < '2021-12-30', btxk.TSXK7_2020,  
	IIF(ut.NgayTK < '2022-06-30', btxk.TSXK2022, btxk.TSXK7_2022))) as TSKhaibao,

	ur.HSKhaibao, ur.HSKiemtra, 

	--IIF(ut.NgayTK between '2018-01-01' and '2020-07-09', btx.TSXK125,
	--IIF( ut.NgayTK between '2020-07-10' and '2021-12-29', btx.TSXK7_2020,  
	--IIF(ut.NgayTK between '2021-12-30' and '2022-06-29', btx.TSXK2022, btx.TSXK7_2022))) as TSPhanloai, 

	IIF(ut.NgayTK < '2020-07-10', btx.TSXK125,
	IIF(ut.NgayTK < '2021-12-30', btx.TSXK7_2020,  
	IIF(ut.NgayTK < '2022-06-30', btx.TSXK2022, btx.TSXK7_2022))) as TSPhanloai,

	ut.MotaHanghoa, ur.MotaHanghoa as TenhangKhaibao, ur.ThongtinRuiro, ut.MaDN, ut.TenDN, ut.TenDoitac, 
    ut.Soluong, ut.DVT, ut.ManuocXX, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT, 
	ur.VanbanThamchieu as SoTBKQPL, ur.NgayVB as NgayTBKQPL, ut.NgaynhapHT, ut.LoginName
	 
    From UserTokhai as ut 
	Inner Join UserDanhmucQLRR as ur On ut.MasoHS = ur.HSKhaibao 
	Inner Join BieuthueXK as btxk On ur.HSKhaibao = btxk.MasoHS 
    Inner Join BieuthueXK as btx On ur.HSKiemtra = btx.MasoHS
	Inner Join DulieuV5 as v5 On ut.SoTK = v5.SoTK 
	Where (ut.LoginName = @utloginName and ur.LoginName = @urloginName) and ur.Trangthai = '1' and Left(ut.SoTK,1) = '3'
END
GO

CREATE PROCEDURE [dbo].[AnalyseRRNK_sp]   
	-- Add the parameters for the stored procedure here
	@utloginName varchar (9), @urloginName varchar (9)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	Select Distinct  ROW_NUMBER() OVER (ORDER BY ut.MasoHS ASC) AS [STT], ur.Trangthai, ut.SoTK, v5.MaHQ, ut.NgayTK, ut.MaLoaihinh, 
	v5.LuongTK, v5.SoVandon, v5.TenCCDKV5, v5.TenCCKHV5, v5.NgayHTKT, v5.NgayGPH, v5.NgayGPHV5, v5.NgayBQ, v5.NgayCPXN, v5.PhanGhichu,
	ut.SttHang, ut.MasoHS, 

	--IIF(ut.NgayTK between '2018-01-01' and '2020-07-09', btxk.TSXK125,
	--IIF( ut.NgayTK between '2020-07-10' and '2021-12-29', btxk.TSXK7_2020,  
	--IIF(ut.NgayTK between '2021-12-30' and '2022-06-29', btxk.TSXK2022, btxk.TSXK7_2022))) as TSKhaibao,

	IIF(ut.NgayTK < '2020-07-10', btnk.TSXK125,
	IIF(ut.NgayTK < '2021-12-30', btnk.TSXK7_2020,  
	IIF(ut.NgayTK < '2022-06-30', btnk.TSXK2022, btxk.TSXK7_2022))) as TSKhaibao,

	ur.HSKhaibao, ur.HSKiemtra, 

	--IIF(ut.NgayTK between '2018-01-01' and '2020-07-09', btx.TSXK125,
	--IIF( ut.NgayTK between '2020-07-10' and '2021-12-29', btx.TSXK7_2020,  
	--IIF(ut.NgayTK between '2021-12-30' and '2022-06-29', btx.TSXK2022, btx.TSXK7_2022))) as TSPhanloai, 

	IIF(ut.NgayTK < '2020-07-10', btn.TSXK125,
	IIF(ut.NgayTK < '2021-12-30', btn.TSXK7_2020,  
	IIF(ut.NgayTK < '2022-06-30', btn.TSXK2022, btn.TSXK7_2022))) as TSPhanloai,

	ut.MotaHanghoa, ur.MotaHanghoa as TenhangKhaibao, ur.ThongtinRuiro, ut.MaDN, ut.TenDN, ut.TenDoitac, 
    ut.Soluong, ut.DVT, ut.ManuocXX, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT, 
	ur.VanbanThamchieu as SoTBKQPL, ur.NgayVB as NgayTBKQPL, ut.NgaynhapHT, ut.LoginName
	 
    From UserTokhai as ut 
	Inner Join UserDanhmucQLRR as ur On ut.MasoHS = ur.HSKhaibao 
	Inner Join BieuthueNK as btnk On ur.HSKhaibao = btnk.MasoHS 
    Inner Join BieuthueNK as btn On ur.HSKiemtra = btn.MasoHS
	Inner Join DulieuV5 as v5 On ut.SoTK = v5.SoTK 
	Where (ut.LoginName = @utloginName and ur.LoginName = @urloginName) and ur.Trangthai = '1' and Left(ut.SoTK,1) = '1' and ur.XN = 'N'
END
GO

Select Distinct  ROW_NUMBER() OVER (ORDER BY ut.MasoHS ASC) AS [STT], uk.Trangthai, ut.SoTK, v5.MaHQ, ut.NgayTK, ut.MaLoaihinh, 
	v5.LuongTK, v5.SoVandon, v5.TenCCDKV5, v5.TenCCKHV5, v5.NgayHTKT, v5.NgayGPH, v5.NgayGPHV5, v5.NgayBQ, v5.NgayCPXN, v5.PhanGhichu,
	ut.SttHang, ut.MasoHS, 
	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btnk.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btnk.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btnk.VJEPA,	
	IIF(ut.ManuocXX = 'KR', btnk.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btnk.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btnk.AHKFTA,  
	IIF(ut.ManuocXX = 'FR' OR ut.ManuocXX = 'DE' OR ut.ManuocXX = 'IT' OR ut.ManuocXX = 'BE' OR ut.ManuocXX = 'NL' OR ut.ManuocXX = 'LU' 
	OR ut.ManuocXX = 'IE' OR ut.ManuocXX = 'DK' OR ut.ManuocXX = 'GR' OR ut.ManuocXX = 'EG' OR ut.ManuocXX = 'ES' OR ut.ManuocXX = 'PT'	
	OR ut.ManuocXX = 'AT' OR ut.ManuocXX = 'SE'OR ut.ManuocXX = 'FI' OR ut.ManuocXX = 'CZ' OR ut.ManuocXX = 'HU' OR ut.ManuocXX = 'PL' 
	OR ut.ManuocXX = 'SK' OR ut.ManuocXX = 'SI' OR ut.ManuocXX = 'LT' OR ut.ManuocXX = 'LV' OR ut.ManuocXX = 'EE' OR ut.ManuocXX = 'MT' 
	OR ut.ManuocXX = 'CY' OR ut.ManuocXX = 'BG' OR ut.ManuocXX = 'RO' OR ut.ManuocXX = 'HR', btnk.EVFTA, btnk.TNKUD))))))) as TSKhaibao, 

	uk.MasoHS as HSKhaibao, uk.MasoPhanloai as HSKiemtra, 

	IIF(ut.ManuocXX = 'CN' OR ut.TennuocXX = 'CN', btn.ACFTA, 
	IIF(ut.ManuocXX = 'BN' OR ut.ManuocXX = 'KH' OR ut.ManuocXX = 'ID' OR ut.ManuocXX = 'LA' 
	OR ut.ManuocXX = 'MY' OR ut.ManuocXX = 'MM' OR ut.ManuocXX = 'PH' OR ut.ManuocXX = 'SG' OR ut.ManuocXX = 'TH', btn.ATIGA, 
	IIF(ut.ManuocXX = 'JP', btn.VJEPA,	
	IIF(ut.ManuocXX = 'KR', btn.VKFTA, 
	IIF(ut.ManuocXX = 'CL', btn.VCFTA, 
	IIF(ut.ManuocXX = 'HK', btn.AHKFTA, 

	IIF(ut.ManuocXX = 'FR' OR ut.ManuocXX = 'DE' OR ut.ManuocXX = 'IT' OR ut.ManuocXX = 'BE' OR ut.ManuocXX = 'NL' OR ut.ManuocXX = 'LU' 
	OR ut.ManuocXX = 'IE' OR ut.ManuocXX = 'DK' OR ut.ManuocXX = 'GR' OR ut.ManuocXX = 'EG' OR ut.ManuocXX = 'ES' OR ut.ManuocXX = 'PT'	
	OR ut.ManuocXX = 'AT' OR ut.ManuocXX = 'SE'OR ut.ManuocXX = 'FI' OR ut.ManuocXX = 'CZ' OR ut.ManuocXX = 'HU' OR ut.ManuocXX = 'PL' 
	OR ut.ManuocXX = 'SK' OR ut.ManuocXX = 'SI' OR ut.ManuocXX = 'LT' OR ut.ManuocXX = 'LV' OR ut.ManuocXX = 'EE' OR ut.ManuocXX = 'MT' 
	OR ut.ManuocXX = 'CY' OR ut.ManuocXX = 'BG' OR ut.ManuocXX = 'RO' OR ut.ManuocXX = 'HR', btn.EVFTA, btn.TNKUD))))))) as TSPhanloai, 

	ut.MotaHanghoa, uk.TenhangKhaibao, uk.MotaHanghoa as ThongtinRuiro, ut.MaDN, ut.TenDN, ut.TenDoitac, 
    ut.Soluong, ut.DVT, ut.ManuocXX, ut.TennuocXX, ut.Dongia, ut.MaNT, ut.TrigiaNT, ut.TrigiaTT, 
	uk.SoTBKQPL, uk.NgayTBKQPL, ut.NgaynhapHT, ut.LoginName
	 
    From UserTokhai as ut Inner Join UserKQPTPL as uk On ut.MasoHS = uk.MasoHS Inner Join BieuthueNK as btnk On uk.MasoHS = btnk.MasoHS 
    Inner Join BieuthueNK as btn On uk.MasoPhanloai = btn.MasoHS
	Inner Join DulieuV5 as v5 On ut.SoTK = v5.SoTK 
	Where (ut.LoginName = 'HQ10-0152' and uk.LoginName = 'HQ10-0152') and uk.Trangthai = '1' and Left(ut.SoTK,1) = '1' --Order By ut.MasoHS ASC