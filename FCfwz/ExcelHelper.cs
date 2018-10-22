﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data;
using System.IO;

namespace FCfwz
{
    public class ExcelHelper
    {
        public static IWorkbook ToExcel(DataTable table, string title)
        {

            IWorkbook workBook = new HSSFWorkbook();
            ISheet sheet = workBook.CreateSheet("Sheet1"); 

            sheet.DefaultRowHeightInPoints = 20;

            HSSFFont fontX = (HSSFFont)workBook.CreateFont();
            fontX.FontName = "宋体";
            fontX.IsBold = true;
            fontX.FontHeightInPoints = 10;

            ICellStyle styleX = workBook.CreateCellStyle();
            //设置单元格上下左右边框线  
            styleX.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styleX.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleX.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styleX.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            //文字水平和垂直对齐方式  
            styleX.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleX.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            //是否换行  
            //cellStyle.WrapText = true;  
            //缩小字体填充  

            styleX.ShrinkToFit = false;
            styleX.SetFont(fontX);

            HSSFFont fontY = (HSSFFont)workBook.CreateFont();
            fontY.FontName = "宋体";
            fontY.FontHeightInPoints = 10;

            ICellStyle styleY = workBook.CreateCellStyle();
            //设置单元格上下左右边框线  
            styleY.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styleY.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleY.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styleY.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            //文字水平和垂直对齐方式  
            styleY.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            styleY.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            //是否换行  
            //cellStyle.WrapText = true;  
            //缩小字体填充  
            styleY.SetFont(fontY);
            styleY.ShrinkToFit = false;


            ICellStyle styleZ = workBook.CreateCellStyle();
            //设置单元格上下左右边框线  
            styleZ.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styleZ.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleZ.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styleZ.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            //文字水平和垂直对齐方式  
            styleZ.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            styleZ.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            //是否换行  
            //cellStyle.WrapText = true;  
            //缩小字体填充  
            styleZ.SetFont(fontY);
            styleZ.ShrinkToFit = false; 


            ICellStyle cellStyleZ = workBook.CreateCellStyle();
            IFont fontZ = workBook.CreateFont();
            fontZ.FontName = "微软雅黑";
            fontZ.FontHeightInPoints = 17;
            cellStyleZ.SetFont(fontZ);
            cellStyleZ.VerticalAlignment = VerticalAlignment.Center;
            cellStyleZ.Alignment = HorizontalAlignment.Center;


            //处理表格标题
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(title);
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, table.Columns.Count - 1));
            row.Height = 500;
            row.Cells[0].CellStyle = cellStyleZ;


            //显示条件
            row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("");
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, table.Columns.Count - 1));
            row.HeightInPoints = 20;//行高 

            //准备表头
            row = sheet.CreateRow(2);
            row.HeightInPoints = 20;//行高  
            for (int i = 0; i < table.Columns.Count; i++)
            {
                row.CreateCell(i).SetCellValue(table.Columns[i].ColumnName);
                row.Cells[i].CellStyle = styleX; 
            }
            row = sheet.CreateRow(3);
            row.HeightInPoints = 20;//行高  
            for (int i = 0; i < table.Columns.Count; i++)
            {
                row.CreateCell(i).SetCellValue(table.Columns[i].ColumnName);
                row.Cells[i].CellStyle = styleX; 
            }

            //处理数据内容
            for (int i = 0; i < table.Rows.Count; i++)
            {
                row = sheet.CreateRow(4 + i);
                row.HeightInPoints = 20;//行高 
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    DataColumn dc = table.Columns[j];
                    if (dc.DataType == System.Type.GetType("System.Decimal"))
                    {
                        row.CreateCell(j).SetCellValue(Convert.ToDecimal(table.Rows[i][j]).ToString("f2"));
                        row.Cells[j].CellStyle = styleZ;
                    }
                    else if (dc.DataType == System.Type.GetType("System.Int32"))
                    {
                        row.CreateCell(j).SetCellValue(Convert.ToDecimal(table.Rows[i][j]).ToString("f0"));
                        row.Cells[j].CellStyle = styleZ; ;
                    }
                    else
                    {
                        row.CreateCell(j).SetCellValue(table.Rows[i][j].ToString());
                        row.Cells[j].CellStyle = styleY;
                    }

                    //sheet.AutoSizeColumn(j);
                }
            }


            return workBook;

            ////创建一个 IO 流
            //MemoryStream ms = new MemoryStream();
            ////写入数据流
            //workBook.Write(ms);
            ////转换为字节数组
            //byte[] bytes = ms.ToArray();
            //ms.Close();
            //return bytes;
        }

    }
}

