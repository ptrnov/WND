if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[st01a]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[st01a]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[st01b]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[st01b]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[st01c]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[st01c]
GO

CREATE TABLE [dbo].[st01a] (
	[USER_ID] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[USR_PASS] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[USR_NM] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[USR_OFF] [tinyint] NOT NULL ,
	[KAR_ID] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DEP_ID] [int] NOT NULL ,
	[JABATAN_ID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[st01b] (
	[MN_ID] [int] NOT NULL ,
	[MN_PRN] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MN_NM] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MN_ORD] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MN_FILE] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MN_OFF] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TMPL_ID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[st01c] (
	[PRMS_ID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MN_ID] [int] NOT NULL ,
	[MN_SHW] [int] NULL ,
	[MN_RD] [int] NULL ,
	[MN_WR] [int] NULL ,
	[MN_DEL] [int] NULL 
) ON [PRIMARY]
GO

