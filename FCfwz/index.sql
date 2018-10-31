if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H4_收款记录]'))
drop index [dbo].[H4_收款记录].[_total]
GO

if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H3_划价记录]'))
drop index [dbo].[H3_划价记录].[_total]
GO

CREATE  INDEX [_total] ON [dbo].[H3_划价记录]([日期], [交款标志], [C3]) ON [PRIMARY]
GO

CREATE  INDEX [_total] ON [dbo].[H4_收款记录]([日期],[发票号]) ON [PRIMARY]
GO
