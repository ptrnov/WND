/*
Navicat MySQL Data Transfer

Source Server         : LocalhostMySQL
Source Server Version : 50532
Source Host           : localhost:3306
Source Database       : menu

Target Server Type    : SQL Server
Target Server Version : 80000
File Encoding         : 65001

Date: 2014-05-06 23:24:10
*/


GO

-- ----------------------------
-- Table structure for [st01a]
-- ----------------------------
DROP TABLE [st01a]
GO
CREATE TABLE [st01a] (
[USER_ID] varchar(15) NOT NULL ,
[USR_PASS] varchar(100) NOT NULL ,
[USR_NM] varchar(200) NOT NULL ,
[USR_OFF] tinyint NOT NULL ,
[KAR_ID] varchar(100) NOT NULL ,
[DEP_ID] int NOT NULL ,
[JABATAN_ID] int NOT NULL 
)


GO

-- ----------------------------
-- Records of st01a
-- ----------------------------
BEGIN TRANSACTION
GO
INSERT INTO [st01a] VALUES (N'ROOT', N'asd123', N'ADMINISTRATOR', N'0', N'001', N'2', N'2');
INSERT INTO [st01a] VALUES (N'endny', N'asd123', N'ENDNY', N'0', N'', N'0', N'0');
INSERT INTO [st01a] VALUES (N'lingling', N'asd123', N'LING LING', N'0', N'', N'0', N'0');
GO
COMMIT TRANSACTION
GO

-- ----------------------------
-- Table structure for [st01b]
-- ----------------------------
DROP TABLE [st01b]
GO
CREATE TABLE [st01b] (
[MN_ID] int NOT NULL ,
[MN_PRN] varchar(5) NOT NULL ,
[MN_NM] varchar(50) NOT NULL ,
[MN_ORD] varchar(10) NOT NULL ,
[MN_FILE] varchar(100) NOT NULL ,
[MN_OFF] varchar(1) NOT NULL ,
[TMPL_ID] int NOT NULL 
)


GO

-- ----------------------------
-- Records of st01b
-- ----------------------------
BEGIN TRANSACTION
GO
INSERT INTO [st01b] VALUES (N'1', N'0', N'Home', N'A', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'2', N'0', N'Basic', N'B', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'3', N'0', N'Transaksi', N'C', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'4', N'0', N'Report', N'D', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'5', N'0', N'Maintenance', N'E', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'6', N'0', N'Admin', N'F', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'50', N'1', N'Informasi ', N'A.1', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'51', N'2', N'Periode', N'B.1', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'52', N'2', N'Setup GL', N'B.2', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'53', N'2', N'Kurs', N'B.3', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'54', N'2', N'Akun', N'B.4', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'55', N'2', N'Cost', N'B.5', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'56', N'3', N'Jurnal Umum', N'C.1', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'57', N'3', N'Posting', N'C.2', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'58', N'3', N'Tutup Tahun', N'C.3', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'59', N'4', N'Analize Report', N'D.1', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'60', N'4', N'Accounting Report ', N'D.2', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'61', N'5', N'Import', N'E.1', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'62', N'6', N'User', N'F.1', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'63', N'6', N'User Permission', N'F.2', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'30', N'30', N'Option', N'', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'31', N'30', N'Swich User', N'', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'32', N'30', N'Log Off', N'', N'', N'', N'0');
INSERT INTO [st01b] VALUES (N'64', N'5', N'Log System', N'E.2', N'', N'', N'0');
GO
COMMIT TRANSACTION
GO

-- ----------------------------
-- Table structure for [st01c]
-- ----------------------------
DROP TABLE [st01c]
GO
CREATE TABLE [st01c] (
[PRMS_ID] varchar(50) NOT NULL ,
[MN_ID] int NOT NULL ,
[MN_SHW] int NOT NULL ,
[MN_RD] int NOT NULL ,
[MN_WR] int NOT NULL ,
[MN_DEL] int NOT NULL 
)


GO

-- ----------------------------
-- Records of st01c
-- ----------------------------
BEGIN TRANSACTION
GO
INSERT INTO [st01c] VALUES (N'ROOT', N'1', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'2', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'3', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'4', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'5', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'6', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'50', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'51', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'52', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'53', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'54', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'55', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'56', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'57', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'58', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'59', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'60', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'61', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'62', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'63', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'ROOT', N'64', N'1', N'1', N'1', N'1');
INSERT INTO [st01c] VALUES (N'lingling', N'1', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'2', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'3', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'4', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'5', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'6', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'50', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'51', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'52', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'53', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'54', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'55', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'56', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'57', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'58', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'59', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'60', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'61', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'62', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'63', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'lingling', N'64', N'1', N'1', N'1', N'1');
INSERT INTO [st01c] VALUES (N'endny', N'1', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'2', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'3', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'4', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'5', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'6', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'50', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'51', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'52', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'53', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'54', N'1', N'1', N'1', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'55', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'56', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'57', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'58', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'59', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'60', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'61', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'62', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'63', N'1', N'1', N'0', N'0');
INSERT INTO [st01c] VALUES (N'endny', N'64', N'1', N'1', N'1', N'1');
GO
COMMIT TRANSACTION
GO
