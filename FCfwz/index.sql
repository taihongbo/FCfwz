if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H4_�տ��¼]'))
drop index [dbo].[H4_�տ��¼].[_total]
GO

if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H3_���ۼ�¼]'))
drop index [dbo].[H3_���ۼ�¼].[_total]
GO

CREATE  INDEX [_total] ON [dbo].[H3_���ۼ�¼]([����], [�����־], [C3]) ON [PRIMARY]
GO

CREATE  INDEX [_total] ON [dbo].[H4_�տ��¼]([����],[��Ʊ��]) ON [PRIMARY]
GO
