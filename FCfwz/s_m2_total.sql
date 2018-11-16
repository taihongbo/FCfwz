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
		set @whe='A.日期>=''' + @a + ''' And A.日期<=''' + @b +''''
	end
else
	begin
		set @whe='A.日期>=''' + @a + ''' And A.日期<=''' + @b +'''And A.部门编码=''' + @c +''''
	end
 
print(@whe)

set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select  t1.类码,t1.类名 , sl_t2 ,je_t2 , t3.药品编码,t3.药品名称,t3.规格,t3.单位,t3.单价,sl_t3,je_t3 ,t4.部门编码,t4.部门名称,sl_t4,je_t4 ,t5.医师编码,t5.医师名称,sl_t5,je_t5 from '
set @sql=@sql + ' ('
set @sql=@sql + ' (Select B.类码, B.类名,sum(A.总量) as sl_t2,sum(A.金额) as je_t2  '
set @sql=@sql + '   From H3_划价记录 As A Left Join  H7_药典总帐 AS B On B.药码 = A.药品编码 '
set @sql=@sql + '   Where A.C3=''门诊'' And  A.交款标志 in (''已交款'',''已发药'')  And '+ @whe 
set @sql=@sql + '   Group By B.类码,B.类名'
set @sql=@sql + ' ) T1 '
set @sql=@sql + ' inner join '
set @sql=@sql + ' (Select B.类码, B.类名, A.药品编码,A.药品名称,A.规格,A.单位,A.单价,sum(A.总量) as sl_t3,sum(A.金额) as je_t3  '
set @sql=@sql + '   From H3_划价记录 As A Left Join  H7_药典总帐 AS B On B.药码 = A.药品编码   '
set @sql=@sql + '   Where A.C3=''门诊'' And  A.交款标志 in (''已交款'',''已发药'')  And '+ @whe 
set @sql=@sql + '   Group By B.类码, B.类名 , A.药品编码,A.药品名称,A.规格,A.单位,A.单价 '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t1.类码 =t3.类码 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select B.类码, B.类名, A.部门编码,A.部门名称,A.药品编码,A.药品名称,A.规格,A.单位,A.单价,sum(A.总量) as sl_t4,sum(A.金额) as je_t4 '
set @sql=@sql + '   From H3_划价记录 As A Left Join  H7_药典总帐 AS B On B.药码 = A.药品编码   '
set @sql=@sql + '   Where A.C3=''门诊'' And  A.交款标志 in (''已交款'',''已发药'')   And '+ @whe  
set @sql=@sql + '   Group By B.类码, B.类名 ,A.部门编码,A.部门名称,A.药品编码,A.药品名称,A.规格,A.单位,A.单价  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t1.类码 =t4.类码 And t3.药品编码 =t4.药品编码 And t3.药品名称 =t4.药品名称 And t3.规格 =t4.规格 And t3.单位 =t4.单位 And t3.单价 =t4.单价  '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select B.类码, B.类名, A.部门编码,A.部门名称,A.医师编码,A.医师名称,A.药品编码,A.药品名称,A.规格,A.单位,A.单价,sum(A.总量) as sl_t5,sum(A.金额) as je_t5 '
set @sql=@sql + '   From H3_划价记录 As A Left Join  H7_药典总帐 AS B On B.药码 = A.药品编码   '
set @sql=@sql + '   Where A.C3=''门诊'' And  A.交款标志 in (''已交款'',''已发药'')    And '+ @whe 
set @sql=@sql + '   Group By B.类码, B.类名, A.部门编码,A.部门名称,A.医师编码,A.医师名称,A.药品编码,A.药品名称,A.规格,A.单位,A.单价  '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On t1.类码 =t5.类码 And t4.部门编码 =t5.部门编码 And t4.部门名称 =t5.部门名称 And t4.药品编码 =t5.药品编码 And t4.药品名称 =t5.药品名称 And t4.规格 =t5.规格 And t4.单位 =t5.单位 And t4.单价 =t5.单价  '
set @sql=@sql + ' )'
set @sql=@sql + ' ) T6 Order By T6.类码 , T6.药品编码 , T6.部门编码 , T6.医师编码' 
print(@sql)
execute(@sql) 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

