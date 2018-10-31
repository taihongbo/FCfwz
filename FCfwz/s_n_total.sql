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
	set @whe='日期>=''' + @a + ''' And 日期<=''' + @b +''''
else
	set @whe='结算日期>=''' + @a + ''' And 结算日期<=''' + @b +'''' 

set @sql=@sql + 'Select * from ('
set @sql=@sql + ' Select  t1.科码,t1.科名,t1.项目编码,t1.项目名称, sl_t3 as 数量,je_t3 as 金额 ,t4.部门编码,t4.部门名称,sl_t4 as 部门数量 ,je_t4 as 部门金额 ,t5.医师编码,t5.医师名称,sl_t5 as 医师数量,je_t5 as 医师金额 from '
if (@c<>'')
	set @sql=@sql + ' (Select 科码,科名,项目编码,项目名称   from H0_收费项目 Where 科码=''' + ltrim(rtrim(@c)) + ''' group by  科码,科名,项目编码,项目名称  ) t1'
else 
	set @sql=@sql + ' (Select 科码,科名,项目编码,项目名称   from H0_收费项目 Where 科码<>'''' group by  科码,科名,项目编码,项目名称  ) t1'

set @sql=@sql + ' inner join '
set @sql=@sql + ' (Select 项目编码, 项目名称,count(H4_收款记录) as sl_t3,sum(金额) as je_t3  '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where '+ @whe 
set @sql=@sql + '   Group By 项目编码, 项目名称 '
set @sql=@sql + ' ) T3 '
set @sql=@sql + '  On t1.项目编码 =t3.项目编码 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 部门编码,部门名称,项目编码, 项目名称,count(H4_收款记录) as sl_t4,sum(金额) as je_t4 '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where  '+ @whe  
set @sql=@sql + '   Group By 部门编码,部门名称,项目编码, 项目名称  '
set @sql=@sql + ' ) T4 '
set @sql=@sql + ' On t3.项目编码 =t4.项目编码 And t3.项目名称 =t4.项目名称 '
set @sql=@sql + '  inner join  '
set @sql=@sql + ' (Select 部门编码,部门名称,医师编码,医师名称,项目编码, 项目名称,count(H4_收款记录) as sl_t5,sum(金额) as je_t5 '
set @sql=@sql + '   From H4_收款记录   '
set @sql=@sql + '   Where  '+ @whe 
set @sql=@sql + '   Group By 部门编码,部门名称,医师编码,医师名称,项目编码,项目名称 '
set @sql=@sql + ' ) T5 '
set @sql=@sql + ' On t4.部门编码 =t5.部门编码 And t4.部门名称 =t5.部门名称 And t4.项目编码 =t5.项目编码 And t4.项目名称 =t5.项目名称 '
set @sql=@sql + ') as T6 Order By T6.科码,T6.项目编码,T6.部门编码,T6.医师编码' 

print(@sql)
execute(@sql)  




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

