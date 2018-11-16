SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
 
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[s_m2_total]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[s_m2_total]
GO 

CREATE PROCEDURE s_m2_total
( 
	@a char(10),
	@b char(10),
    @c char(10)
)
AS

declare @sql varchar(8000) 
set @sql=''


declare @whe varchar(1000)
set @whe=''
set @a=ltrim(rtrim(@a))
set @b=ltrim(rtrim(@b))
set @c=ltrim(rtrim(@c))

if (@c='')
	begin
		set @whe='A.����>=''' + @a + ''' And A.����<=''' + @b +''''
	end
else
	begin
		set @whe='A.����>=''' + @a + ''' And A.����<=''' + @b +'''And A.���ű���=''' + @c +''''
	end
 
print(@whe)

set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select  t1.����,t1.���� , sl_t2 ,je_t2 , t3.ҩƷ����,t3.ҩƷ����,t3.���,t3.��λ,t3.����,sl_t3,je_t3 ,t4.���ű���,t4.��������,sl_t4,je_t4 ,t5.ҽʦ����,t5.ҽʦ����,sl_t5,je_t5 from '
set @sql=@sql + ' ('
set @sql=@sql + ' (Select B.����, B.����,sum(A.����) as sl_t2,sum(A.���) as je_t2  '
set @sql=@sql + '   From H3_���ۼ�¼ As A Left Join  H7_ҩ������ AS B On B.ҩ�� = A.ҩƷ���� '
set @sql=@sql + '   Where A.C3=''����'' And  A.�����־ in (''�ѽ���'',''�ѷ�ҩ'')  And '+ @whe 
set @sql=@sql + '   Group By B.����,B.����'
set @sql=@sql + ' ) T1 '
set @sql=@sql + ' inner join '
set @sql=@sql + ' (Select B.����, B.����, A.ҩƷ����,A.ҩƷ����,A.���,A.��λ,A.����,sum(A.����) as sl_t3,sum(A.���) as je_t3  '
set @sql=@sql + '   From H3_���ۼ�¼ As A Left Join  H7_ҩ������ AS B On B.ҩ�� = A.ҩƷ����   '
set @sql=@sql + '   Where A.C3=''����'' And  A.�����־ in (''�ѽ���'',''�ѷ�ҩ'')  And '+ @whe 
set @sql=@sql + '   Group By B.����, B.���� , A.ҩƷ����,A.ҩƷ����,A.���,A.��λ,A.���� '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t1.���� =t3.���� '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select B.����, B.����, A.���ű���,A.��������,A.ҩƷ����,A.ҩƷ����,A.���,A.��λ,A.����,sum(A.����) as sl_t4,sum(A.���) as je_t4 '
set @sql=@sql + '   From H3_���ۼ�¼ As A Left Join  H7_ҩ������ AS B On B.ҩ�� = A.ҩƷ����   '
set @sql=@sql + '   Where A.C3=''����'' And  A.�����־ in (''�ѽ���'',''�ѷ�ҩ'')   And '+ @whe  
set @sql=@sql + '   Group By B.����, B.���� ,A.���ű���,A.��������,A.ҩƷ����,A.ҩƷ����,A.���,A.��λ,A.����  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t1.���� =t4.���� And t3.ҩƷ���� =t4.ҩƷ���� And t3.ҩƷ���� =t4.ҩƷ���� And t3.��� =t4.��� And t3.��λ =t4.��λ And t3.���� =t4.����  '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select B.����, B.����, A.���ű���,A.��������,A.ҽʦ����,A.ҽʦ����,A.ҩƷ����,A.ҩƷ����,A.���,A.��λ,A.����,sum(A.����) as sl_t5,sum(A.���) as je_t5 '
set @sql=@sql + '   From H3_���ۼ�¼ As A Left Join  H7_ҩ������ AS B On B.ҩ�� = A.ҩƷ����   '
set @sql=@sql + '   Where A.C3=''����'' And  A.�����־ in (''�ѽ���'',''�ѷ�ҩ'')    And '+ @whe 
set @sql=@sql + '   Group By B.����, B.����, A.���ű���,A.��������,A.ҽʦ����,A.ҽʦ����,A.ҩƷ����,A.ҩƷ����,A.���,A.��λ,A.����  '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On t1.���� =t5.���� And t4.���ű��� =t5.���ű��� And t4.�������� =t5.�������� And t4.ҩƷ���� =t5.ҩƷ���� And t4.ҩƷ���� =t5.ҩƷ���� And t4.��� =t5.��� And t4.��λ =t5.��λ And t4.���� =t5.����  '
set @sql=@sql + ' )'
set @sql=@sql + ' ) T6 Order By T6.���� , T6.ҩƷ���� , T6.���ű��� , T6.ҽʦ����' 
print(@sql)
execute(@sql) 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

