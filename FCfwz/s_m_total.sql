SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[s_m_total]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[s_m_total]
GO

CREATE     PROCEDURE s_m_total
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
set @t=ltrim(rtrim(@t))
set @a=ltrim(rtrim(@a))
set @b=ltrim(rtrim(@b))
set @c=ltrim(rtrim(@c))

if (@c='')
	begin
		if (@t='1') 
			set @whe='����>=''' + @a + ''' And ����<=''' + @b +''''
		else
			set @whe='��������>=''' + @a + ''' And ��������<=''' + @b +''''
	end
else
	begin
		if (@t='1') 
			set @whe='����>=''' + @a + ''' And ����<=''' + @b +'''And ���ű���=''' + @c +''''
		else
			set @whe='��������>=''' + @a + ''' And ��������<=''' + @b +''' And ���ű���=''' + @c +''''
	end

print(@whe)


set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select   t1.����,t1.����,t3.ҩƷ����,t3.ҩƷ����,t3.���,t3.��λ,t3.����,sl_t3 as ����,je_t3 as ��� ,t4.���ű���,t4.��������,sl_t4 as �������� ,je_t4 as ���Ž�� ,t5.ҽʦ����,t5.ҽʦ����,sl_t5 as ҽʦ����,je_t5 as ҽʦ��� from '
set @sql=@sql + ' (Select ����,����   from H7_ҩ����� group by  ����,����  ) t1'
set @sql=@sql + ' inner join '
set @sql=@sql + ' (Select ������, ҩƷ����,ҩƷ����,���,��λ,����,sum(����) as sl_t3,sum(���) as je_t3  '
set @sql=@sql + '   From H3_���ۼ�¼   '
set @sql=@sql + '   Where C3=''����'' And  �����־ in (''�ѽ���'',''�ѷ�ҩ'')  And '+ @whe 
set @sql=@sql + '   Group By ������,ҩƷ����,ҩƷ����,���,��λ,���� '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t1.���� =t3.������ '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select ���ű���,��������,ҩƷ����,ҩƷ����,���,��λ,����,sum(����) as sl_t4,sum(���) as je_t4 '
set @sql=@sql + '   From H3_���ۼ�¼   '
set @sql=@sql + '   Where C3=''����'' And  �����־ in (''�ѽ���'',''�ѷ�ҩ'')   And '+ @whe  
set @sql=@sql + '   Group By ���ű���,��������,ҩƷ����,ҩƷ����,���,��λ,����  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t3.ҩƷ���� =t4.ҩƷ���� And t3.ҩƷ���� =t4.ҩƷ���� And t3.��� =t4.��� And t3.��λ =t4.��λ And t3.���� =t4.����  '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select ���ű���,��������,ҽʦ����,ҽʦ����,ҩƷ����,ҩƷ����,���,��λ,����,sum(����) as sl_t5,sum(���) as je_t5 '
set @sql=@sql + '   From H3_���ۼ�¼   '
set @sql=@sql + '   Where C3=''����'' And  �����־ in (''�ѽ���'',''�ѷ�ҩ'')    And '+ @whe 
set @sql=@sql + '   Group By ���ű���,��������,ҽʦ����,ҽʦ����,ҩƷ����,ҩƷ����,���,��λ,����  '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On t4.���ű��� =t5.���ű��� And t4.�������� =t5.�������� And t4.ҩƷ���� =t5.ҩƷ���� And t4.ҩƷ���� =t5.ҩƷ���� And t4.��� =t5.��� And t4.��λ =t5.��λ And t4.���� =t5.����  '
set @sql=@sql + ') as T6 Order By T6.����,T6.ҩƷ����,T6.���ű���,T6.ҽʦ����' 

print(@sql)
execute(@sql) 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

