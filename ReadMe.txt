1. Processing t_edif = YES

1.Boiler Trigger
==============
CREATE TRIGGER ISGEC_PropertyIU
   ON  Property 
   AFTER  INSERT,UPDATE
AS 
BEGIN
    SET NOCOUNT ON;
    MERGE INTO ISGEC_Property as Target
    USING inserted as Source 
    ON (TARGET.PropertyID = SOURCE.PropertyID and Source.PropertyDefID in (113,123)) 
    WHEN MATCHED  THEN 
    UPDATE SET TARGET.PropertyDefID = SOURCE.PropertyDefID,
	           Target.EntityID = Source.EntityID,
	           Target.Value = convert(nvarchar(50),source.value)
    WHEN NOT MATCHED BY TARGET THEN 
    INSERT (PropertyID, PropertyDefID,EntityID, Value) 
    VALUES (SOURCE.PropertyID,SOURCE.PropertyDefID, Source.EntityID, convert(nvarchar(50),source.Value));

END
GO
2. Trigger
==========
CREATE TRIGGER ISGEC_PropertyD
   ON  Property 
   AFTER  DELETE
AS 
BEGIN
    SET NOCOUNT ON;
    DELETE FROM ISGEC_Property where PropertyID in (select PropertyID from deleted)

END
GO
3. Create TABLE
===============
USE [SMD]
GO

/****** Object:  Table [dbo].[ISGEC_Property]    Script Date: 15/06/2018 14:57:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ISGEC_Property](
	[PropertyID] [bigint] NOT NULL,
	[PropertyDefID] [bigint] NOT NULL,
	[EntityID] [bigint] NOT NULL,
	[Value] [nvarchar](50) NULL,
 CONSTRAINT [PK_ISGEC_Property] PRIMARY KEY CLUSTERED 
(
	[PropertyID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

To Check in ERP
===============
select * from tdmisg131200 where t_edif=1

select * from ttcisg132200 where t_hndl='TRANSMITTALLINES_200'


