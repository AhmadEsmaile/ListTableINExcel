using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace ListTableTOExcel
{
    public class ListTableINExcel
    {
        public void CreateExcel(List<TableModel> model)
        {
            try
            {

                var tableGroup = model
                        .GroupBy(item => new {  item.Table })
                        .Select(group => new
                        {
                           
                            Table = group.Key.Table,
                            Items = group.ToList()
                        }).ToList();


                var dbName = model.GroupBy(item => item.DBName,
                    (key, group) => new { DBName = key, Items = group.ToList() }).ToList();

            

                byte[] bytes = null;
                using (var excelPackage = new ExcelPackage())
                {
                   
                        ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(dbName[0].DBName);

                        sheet.View.RightToLeft = true;
                    var rowIndex = 2;
                    foreach (var table in tableGroup)
                    {
                        var tableName = table.Items[0].Schema + "." + table.Table;

                        sheet.Cells[1, 1].Value = dbName[0].DBName;
                        var colIndex = 1;
                         rowIndex++;
                        sheet.Cells[rowIndex, 1, rowIndex, 5].Merge = true;
                        sheet.Cells[rowIndex, 1, rowIndex, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[rowIndex, 1, rowIndex, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.YellowGreen);

                        sheet.Cells[rowIndex++, colIndex].Value = tableName;
                        sheet.Cells[rowIndex, colIndex++].Value = "نام ستون در جدول";
                        sheet.Cells[rowIndex, colIndex++].Value = "عنوان فارسی";
                        sheet.Cells[rowIndex, colIndex++].Value = "نوع داده";
                        sheet.Cells[rowIndex, colIndex++].Value = "اجباری؟";
                        sheet.Cells[rowIndex, colIndex].Value = "توضیحات";



                        sheet.Cells[1, 1, 2, colIndex].Merge = true;
                        sheet.Cells[1, 1, 1, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[1, 1, 1, colIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PowderBlue);

                        

                        foreach (var reportItem in table.Items)
                        {
                            colIndex = 1;
                            rowIndex++;

                            sheet.Cells[rowIndex, colIndex++].Value = reportItem.Column;
                            sheet.Cells[rowIndex, colIndex++].Value = reportItem.Description;
                            sheet.Cells[rowIndex, colIndex++].Value = reportItem.DataType;
                            sheet.Cells[rowIndex, colIndex++].Value = reportItem.Nullable == "NO" ? "✓" : null;
                            sheet.Cells[rowIndex, colIndex].Value = null;
                        }



                        using (ExcelRange range = sheet.Cells)
                        {
                            range.Style.Font.SetFromFont(new System.Drawing.Font("Tahoma", 9));
                            range.Style.Font.Bold = true;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.ReadingOrder = ExcelReadingOrder.RightToLeft;
                        }

                        const double minWidth = 0.00;
                        const double maxWidth = 50.00;
                        sheet.Cells.AutoFitColumns(minWidth, maxWidth);
                    }

                    bytes = excelPackage.GetAsByteArray();

                }
                string outputPath = @"C:\_AppPro\ListTable\ " + dbName[0].DBName + " .xlsx";
                File.WriteAllBytes(outputPath, bytes);
            }
            catch (Exception e)
            {
                    throw e;
            }



        }
    }
}

