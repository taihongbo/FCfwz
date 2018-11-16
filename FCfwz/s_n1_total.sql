SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[s_n1_total]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[s_n1_total]
GO


CREATE PROCEDURE s_n1_total
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



if (@c='')
	begin
		set @whe='日期>=''' + @a + ''' And 日期<=''' + @b +''' '
	end
else
	begin
		set @whe='日期>=''' + @a + ''' And 日期<=''' + @b +'''And 科码=''' + @c +''''
	end


set @sql=''


set @sql=' Update H4_收款记录 Set 科码='''',科名='''' Where 日期>=''' + @a + ''' And 日期<=''' + @b +''' '
print(@sql)
execute(@sql)  



set @sql=' Update H4_收款记录 Set  H4_收款记录.科码= H0_收费项目.科码, H4_收款记录.科名= H0_收费项目.科名  From H0_收费项目 Where H0_收费项目.项目编码=H4_收款记录.项目编码 And H4_收款记录.科码='''' And  H4_收款记录.日期>=''' + @a + ''' And H4_收款记录.日期<=''' + @b +''' '
print(@sql)
execute(@sql) 

set @sql=' Update H4_收款记录 Set  科码= 部门编码,  科名= 部门名称 Where 科码='''' And  日期>=''' + @a + ''' And 日期<=''' + @b +''' '
print(@sql)
execute(@sql) 

set @sql=''
set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select  t2.科码,t2.科名, sl_t2 as 科码数量,je_t2 as 科码金额 , t3.项目编码,t3.项目名称, sl_t3 as 项目数量,je_t3 as 项目金额 ,t4.部门编码,t4.部门名称,sl_t4 as 部门数量 ,je_t4 as 部门金额 ,t5.医师编码,t5.医师名称,sl_t5 as 医师数量,je_t5 as 医师金额 from '
set @sql=@sql + ' (  '
set @sql=@sql + ' (Select 科码, 科名,count(H4_收款记录) as sl_t2 , sum(金额) as je_t2  '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where '+ @whe 
set @sql=@sql + '   Group By 科码, 科名 '
set @sql=@sql + ' ) T2 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 科码, 科名 , 项目编码, 项目名称,count(H4_收款记录) as sl_t3,sum(金额) as je_t3  '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where '+ @whe 
set @sql=@sql + '   Group By 科码, 科名 , 项目编码, 项目名称 '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t2.科码 =t3.科码 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 科码, 科名 ,部门编码,部门名称,项目编码, 项目名称,count(H4_收款记录) as sl_t4,sum(金额) as je_t4 '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where  '+ @whe  
set @sql=@sql + '   Group By 科码, 科名 ,部门编码,部门名称,项目编码, 项目名称  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t3.科码 = t4.科码 And t3.项目编码 =t4.项目编码 And t3.项目名称 =t4.项目名称 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 科码, 科名 ,部门编码,部门名称,医师编码,医师名称,项目编码, 项目名称,count(H4_收款记录) as sl_t5,sum(金额) as je_t5 '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where  '+ @whe 
set @sql=@sql + '   Group By 科码, 科名 ,部门编码,部门名称,医师编码,医师名称,项目编码,项目名称 '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On  t4.科码 =t5.科码 And t4.部门编码 =t5.部门编码 And t4.部门名称 =t5.部门名称 And t4.项目编码 =t5.项目编码 And t4.项目名称 =t5.项目名称 '
set @sql=@sql + ')'
set @sql=@sql + ') T6 Order By T6.科码,T6.项目编码,T6.部门编码,T6.医师编码' 

print(@sql)
execute(@sql)  




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

