if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H4_收款记录]'))
drop index [dbo].[H4_收款记录].[_total]
GO

if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H3_划价记录]'))
drop index [dbo].[H3_划价记录].[_total]
GO

CREATE  INDEX [_total] ON [dbo].[H3_划价记录]([日期], [交款标志], [C3])  
GO

CREATE  INDEX [_total] ON [dbo].[H4_收款记录]([日期], [部门编码], [科码]) 
GO

CREATE  INDEX [_srcode] ON [dbo].[H4_收款记录]([收入编码]) 
GO

CREATE  INDEX [_riqi] ON [dbo].[H4_收款记录]([日期]) 
GO

Alter table H4_收款记录 add 收入编码 varchar(20) null

Alter table H4_收款记录 drop column 收入编码 