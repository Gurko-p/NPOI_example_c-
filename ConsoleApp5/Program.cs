using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp5
{
    class Program
    {
        const int COL_WIDTH = 450;
        static void Main(string[] args)
        {
            
            try
            {
                HSSFWorkbook wb = new HSSFWorkbook();
                ISheet sheet = wb.CreateSheet("Мой лист!");
                ISheet sheet2 = wb.CreateSheet("Мой лист2!");
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                string path = Path.Combine(Environment.CurrentDirectory, "test.xls");

                IFont font = wb.CreateFont();
                font.IsBold = true;
                font.FontHeightInPoints = 15;
                font.FontName = "Times New Roman";
                font.IsItalic = true;
                ICellStyle boldStyle = wb.CreateCellStyle();
                boldStyle.SetFont(font);

                cell.SetCellValue("Hello");
                cell.CellStyle = boldStyle;
                cell = row.CreateCell(1);
                cell.SetCellValue("World");
                cell.CellStyle = boldStyle;

                IRow row1 = sheet.CreateRow(1);
                ICell cell3 = row1.CreateCell(3);
                cell3.SetCellValue("sadsadasd");

                ICell cellTime = row.CreateCell(3);
                //cellTime.SetCellValue(DateTime.Today);

                ICellStyle cs = wb.CreateCellStyle();
                cs.WrapText = false;
                cs.VerticalAlignment = VerticalAlignment.Center;
                cs.Alignment = HorizontalAlignment.Center;
                cs.BorderBottom = BorderStyle.Medium;
                cs.FillForegroundColor = IndexedColors.Yellow.Index;
                cs.FillPattern = FillPattern.SolidForeground;
                
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 3, 5));
                //sheet.AddMergedRegion(CellRangeAddress.ValueOf("D1:F1"));
                sheet.SetZoom(75);
                

                IFont fontTime = wb.CreateFont();
                fontTime.IsBold = true;
                fontTime.FontHeightInPoints = 15;
                fontTime.FontName = "Arial";
                fontTime.IsItalic = true;
                fontTime.Color = IndexedColors.Teal.Index;
                ICellStyle fontTimeStyle = wb.CreateCellStyle();
                fontTimeStyle.SetFont(fontTime);
                fontTimeStyle.WrapText = false;
                fontTimeStyle.VerticalAlignment = VerticalAlignment.Center;
                fontTimeStyle.Alignment = HorizontalAlignment.Center;
                fontTimeStyle.BorderBottom = BorderStyle.Thin;
                fontTimeStyle.BorderLeft = BorderStyle.Thin;
                fontTimeStyle.BorderRight = BorderStyle.Thin;
                fontTimeStyle.BorderTop = BorderStyle.Thin;
                fontTimeStyle.FillForegroundColor = IndexedColors.Yellow.Index;
                fontTimeStyle.FillPattern = FillPattern.SolidForeground;

                cellTime.CellStyle = fontTimeStyle;
                cellTime.SetCellValue(DateTime.Today.ToShortDateString());

                sheet.SetAutoFilter(new CellRangeAddress(0, sheet.PhysicalNumberOfRows - 1, 0, sheet.GetRow(0).PhysicalNumberOfCells - 1));

                //List<int> autosizeColList = new List<int> { 0, 1, 2, 3, 4, 5 };
                //autosizeColList.ForEach(x => sheet.AutoSizeColumn(x));

                sheet.SetColumnWidth(1, COL_WIDTH * 10);

                using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(stream);
                }
                Console.WriteLine(Path.Combine(Environment.CurrentDirectory, "test.xls"));
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        public static MyCell createCell(HSSFWorkbook wb, IRow row, int column, HorizontalAlignment align)
        {
            ICell cell = row.CreateCell(column);
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.Alignment = align;
            cell.CellStyle = cellStyle;
            return new MyCell { Row = row.RowNum, Column = column };
        }

        public class MyCell
        {
            public int Row { get; set; }
            public int Column { get; set; }
        }
    }
}
