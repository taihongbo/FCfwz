if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H4_�տ��¼]'))
drop index [dbo].[H4_�տ��¼].[_total]
GO

if exists (select * from dbo.sysindexes where name = N'_total' and id = object_id(N'[dbo].[H3_���ۼ�¼]'))
drop index [dbo].[H3_���ۼ�¼].[_total]
GO

CREATE  INDEX [_total] ON [dbo].[H3_���ۼ�¼]([����], [�����־], [C3])  
GO

CREATE  INDEX [_total] ON [dbo].[H4_�տ��¼]([����], [���ű���], [����]) 
GO

CREATE  INDEX [_srcode] ON [dbo].[H4_�տ��¼]([�������]) 
GO

CREATE  INDEX [_riqi] ON [dbo].[H4_�տ��¼]([����]) 
GO

Alter table H4_�տ��¼ add ������� varchar(20) null

Alter table H4_�տ��¼ drop column ������� 