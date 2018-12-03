SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[s_o1_total]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[s_o1_total]
GO


CREATE PROCEDURE s_o1_total
( 
	@t int=0,
	@a char(10),
	@b char(10),
    @c char(10)
)
AS

declare @sql varchar(8000) 
set @sql=''




declare @whe varchar(1000)
set @whe='' 



if (@c='')
	begin
		set @whe='����>=''' + @a + ''' And ����<=''' + @b +''' '
	end
else
	begin
		set @whe='����>=''' + @a + ''' And ����<=''' + @b +'''And ����=''' + @c +''''
	end


set @sql=''


set @sql=' Update H4_�տ��¼ Set ����='''',����='''' Where ����>=''' + @a + ''' And ����<=''' + @b +''' '
print(@sql)
execute(@sql)   

set @sql=' Update H4_�տ��¼ Set  H4_�տ��¼.�������= H0_�շ���Ŀ.������� From H0_�շ���Ŀ Where H0_�շ���Ŀ.��Ŀ����=H4_�տ��¼.��Ŀ���� And  H4_�տ��¼.����>=''' + @a + ''' And H4_�տ��¼.����<=''' + @b +''' '
print(@sql)
execute(@sql)  

set @sql=' Update H4_�տ��¼ Set  H4_�տ��¼.����= H0_�շ���Ŀ.����, H4_�տ��¼.����= H0_�շ���Ŀ.����  From H0_�շ���Ŀ Where H0_�շ���Ŀ.��Ŀ����=H4_�տ��¼.��Ŀ���� And H4_�տ��¼.����='''' And  H4_�տ��¼.����>=''' + @a + ''' And H4_�տ��¼.����<=''' + @b +''' '
print(@sql)
execute(@sql) 

set @sql=' Update H4_�տ��¼ Set  ����= ���ű���,  ����= �������� Where ����='''' And  ����>=''' + @a + ''' And ����<=''' + @b +''' '
print(@sql)
execute(@sql) 

set @sql=''


if (@t=0)
	begin
		set @sql=@sql + 'Select * from ('
		set @sql=@sql + ' Select  t2.����,t2.����, sl_t2 as ��������,je_t2 as ������ , t3.�������,t3.��������, sl_t3 as ��Ŀ����,je_t3 as ��Ŀ��� ,t4.���ű���,t4.��������,sl_t4 as �������� ,je_t4 as ���Ž�� ,t5.ҽʦ����,t5.ҽʦ����,sl_t5 as ҽʦ����,je_t5 as ҽʦ��� from '
		set @sql=@sql + ' (  '
		set @sql=@sql + ' (Select ����, ����,count(H4_�տ��¼) as sl_t2 , sum(���) as je_t2  '
		set @sql=@sql + '   From H4_�տ��¼   '
		set @sql=@sql + '   Where '+ @whe 
		set @sql=@sql + '   Group By ����, ���� '
		set @sql=@sql + ' ) T2 '
		set @sql=@sql + '  inner join  '
		set @sql=@sql + ' (Select ����, ���� , �������, ��������,count(H4_�տ��¼) as sl_t3,sum(���) as je_t3  '
		set @sql=@sql + '   From H4_�տ��¼   '
		set @sql=@sql + '   Where '+ @whe 
		set @sql=@sql + '   Group By ����, ���� , �������, �������� '
		set @sql=@sql + ' ) T3 '
		set @sql=@sql + '  On t2.���� =t3.���� '
		set @sql=@sql + '  inner join  '
		set @sql=@sql + ' (Select ����, ���� ,���ű���,��������,�������, ��������,count(H4_�տ��¼) as sl_t4,sum(���) as je_t4 '
		set @sql=@sql + '   From H4_�տ��¼   '
		set @sql=@sql + '   Where  '+ @whe  
		set @sql=@sql + '   Group By ����, ���� ,���ű���,��������,�������, ��������  '
		set @sql=@sql + ' ) T4 '
		set @sql=@sql + ' On t3.���� = t4.���� And t3.������� =t4.������� And t3.�������� =t4.�������� '
		set @sql=@sql + '  inner join  '
		set @sql=@sql + ' (Select ����, ���� ,���ű���,��������,ҽʦ����,ҽʦ����,�������, ��������,count(H4_�տ��¼) as sl_t5,sum(���) as je_t5 '
		set @sql=@sql + '   From H4_�տ��¼   '
		set @sql=@sql + '   Where  '+ @whe 
		set @sql=@sql + '   Group By ����, ���� ,���ű���,��������,ҽʦ����,ҽʦ����,�������,�������� '
		set @sql=@sql + ' ) T5 '
		set @sql=@sql + ' On  t4.���� =t5.���� And t4.���ű��� =t5.���ű��� And t4.�������� =t5.�������� And t4.������� =t5.������� And t4.�������� =t5.�������� '
		set @sql=@sql + ')'
		set @sql=@sql + ') T6 Order By T6.����,T6.�������,T6.���ű���,T6.ҽʦ����' 

	end
else
	begin
		set @sql=@sql + 'Select * from ('
		set @sql=@sql + ' Select  t2.����,t2.����, sl_t2 as ��������,je_t2 as ������ , t3.�������,t3.��������, sl_t3 as ��Ŀ����,je_t3 as ��Ŀ���  from '
		set @sql=@sql + ' (  '
		set @sql=@sql + ' (Select ����, ����,count(H4_�տ��¼) as sl_t2 , sum(���) as je_t2  '
		set @sql=@sql + '   From H4_�տ��¼   '
		set @sql=@sql + '   Where '+ @whe 
		set @sql=@sql + '   Group By ����, ���� '
		set @sql=@sql + ' ) T2 '
		set @sql=@sql + '  inner join  '
		set @sql=@sql + ' (Select ����, ���� , �������, ��������,count(H4_�տ��¼) as sl_t3,sum(���) as je_t3  '
		set @sql=@sql + '   From H4_�տ��¼   '
		set @sql=@sql + '   Where '+ @whe 
		set @sql=@sql + '   Group By ����, ���� , �������, �������� '
		set @sql=@sql + ' ) T3 '
		set @sql=@sql + '  On t2.���� =t3.���� '
		set @sql=@sql + ')'
		set @sql=@sql + ') T6 Order By T6.����,T6.������� ' 
	end




print(@sql)
execute(@sql)  




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

