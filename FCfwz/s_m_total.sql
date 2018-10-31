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
			set @whe='日期>=''' + @a + ''' And 日期<=''' + @b +''''
		else
			set @whe='交款日期>=''' + @a + ''' And 交款日期<=''' + @b +''''
	end
else
	begin
		if (@t='1') 
			set @whe='日期>=''' + @a + ''' And 日期<=''' + @b +'''And 部门编码=''' + @c +''''
		else
			set @whe='交款日期>=''' + @a + ''' And 交款日期<=''' + @b +''' And 部门编码=''' + @c +''''
	end

print(@whe)


set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select   t1.类码,t1.类名,t3.药品编码,t3.药品名称,t3.规格,t3.单位,t3.单价,sl_t3 as 数量,je_t3 as 金额 ,t4.部门编码,t4.部门名称,sl_t4 as 部门数量 ,je_t4 as 部门金额 ,t5.医师编码,t5.医师名称,sl_t5 as 医师数量,je_t5 as 医师金额 from '
set @sql=@sql + ' (Select 类码,类名   from H7_药典类别 group by  类码,类名  ) t1'
set @sql=@sql + ' inner join '
set @sql=@sql + ' (Select 类别编码, 药品编码,药品名称,规格,单位,单价,sum(总量) as sl_t3,sum(金额) as je_t3  '
set @sql=@sql + '   From H3_划价记录   '
set @sql=@sql + '   Where C3=''门诊'' And  交款标志 in (''已交款'',''已发药'')  And '+ @whe 
set @sql=@sql + '   Group By 类别编码,药品编码,药品名称,规格,单位,单价 '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t1.类码 =t3.类别编码 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 部门编码,部门名称,药品编码,药品名称,规格,单位,单价,sum(总量) as sl_t4,sum(金额) as je_t4 '
set @sql=@sql + '   From H3_划价记录   '
set @sql=@sql + '   Where C3=''门诊'' And  交款标志 in (''已交款'',''已发药'')   And '+ @whe  
set @sql=@sql + '   Group By 部门编码,部门名称,药品编码,药品名称,规格,单位,单价  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t3.药品编码 =t4.药品编码 And t3.药品名称 =t4.药品名称 And t3.规格 =t4.规格 And t3.单位 =t4.单位 And t3.单价 =t4.单价  '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 部门编码,部门名称,医师编码,医师名称,药品编码,药品名称,规格,单位,单价,sum(总量) as sl_t5,sum(金额) as je_t5 '
set @sql=@sql + '   From H3_划价记录   '
set @sql=@sql + '   Where C3=''门诊'' And  交款标志 in (''已交款'',''已发药'')    And '+ @whe 
set @sql=@sql + '   Group By 部门编码,部门名称,医师编码,医师名称,药品编码,药品名称,规格,单位,单价  '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On t4.部门编码 =t5.部门编码 And t4.部门名称 =t5.部门名称 And t4.药品编码 =t5.药品编码 And t4.药品名称 =t5.药品名称 And t4.规格 =t5.规格 And t4.单位 =t5.单位 And t4.单价 =t5.单价  '
set @sql=@sql + ') as T6 Order By T6.类码,T6.药品编码,T6.部门编码,T6.医师编码' 

print(@sql)
execute(@sql) 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

