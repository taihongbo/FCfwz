SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[s_n_total]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[s_n_total]
GO


CREATE       PROCEDURE s_n_total
( 
	@t char(1),
	@a char(10),
	@b char(10),
        @c char(10)
)
AS

declare @sql varchar(8000) 
set @sql=''


declare @whe varchar(1000)
set @whe='' 

if (@t='1') 
	set @whe='����>=''' + @a + ''' And ����<=''' + @b +''''
else
	set @whe='��������>=''' + @a + ''' And ��������<=''' + @b +'''' 

set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select  t1.����,t1.����,t1.��Ŀ����,t1.��Ŀ����, sl_t3 as ����,je_t3 as ��� ,t4.���ű���,t4.��������,sl_t4 as �������� ,je_t4 as ���Ž�� ,t5.ҽʦ����,t5.ҽʦ����,sl_t5 as ҽʦ����,je_t5 as ҽʦ��� from '
if (@c<>'')
	set @sql=@sql + ' (Select ����,����,��Ŀ����,��Ŀ����   from H0_�շ���Ŀ Where ����=''' + ltrim(rtrim(@c)) + ''' group by  ����,����,��Ŀ����,��Ŀ����  ) t1'
else 
	set @sql=@sql + ' (Select ����,����,��Ŀ����,��Ŀ����   from H0_�շ���Ŀ Where ����<>'''' group by  ����,����,��Ŀ����,��Ŀ����  ) t1'

set @sql=@sql + ' inner join '
set @sql=@sql + ' (Select ��Ŀ����, ��Ŀ����,count(H4_�տ��¼) as sl_t3,sum(���) as je_t3  '
set @sql=@sql + '   From H4_�տ��¼   '
set @sql=@sql + '   Where '+ @whe 
set @sql=@sql + '   Group By ��Ŀ����, ��Ŀ���� '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t1.��Ŀ���� =t3.��Ŀ���� '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select ���ű���,��������,��Ŀ����, ��Ŀ����,count(H4_�տ��¼) as sl_t4,sum(���) as je_t4 '
set @sql=@sql + '   From H4_�տ��¼   '
set @sql=@sql + '   Where  '+ @whe  
set @sql=@sql + '   Group By ���ű���,��������,��Ŀ����, ��Ŀ����  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t3.��Ŀ���� =t4.��Ŀ���� And t3.��Ŀ���� =t4.��Ŀ���� '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select ���ű���,��������,ҽʦ����,ҽʦ����,��Ŀ����, ��Ŀ����,count(H4_�տ��¼) as sl_t5,sum(���) as je_t5 '
set @sql=@sql + '   From H4_�տ��¼   '
set @sql=@sql + '   Where  '+ @whe 
set @sql=@sql + '   Group By ���ű���,��������,ҽʦ����,ҽʦ����,��Ŀ����,��Ŀ���� '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On t4.���ű��� =t5.���ű��� And t4.�������� =t5.�������� And t4.��Ŀ���� =t5.��Ŀ���� And t4.��Ŀ���� =t5.��Ŀ���� '
set @sql=@sql + ') as T6 Order By T6.����,T6.��Ŀ����,T6.���ű���,T6.ҽʦ����' 

print(@sql)
execute(@sql)  




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

