using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelTableFormatter
    {
        private static readonly int tableStartColumn = 4;

        public void Format(ClosedXML.Excel.IXLWorksheet worksheet, Parser.Table table, ref int row)
        {
            int startRow = row;
            int headerColumn = tableStartColumn;
            foreach (var cell in table.HeaderRow)
            {
                worksheet.Cell(row, headerColumn++).Value = cell;
            }
            row++;

            foreach (var dataRow in table.DataRows)
            {
                int dataColumn = tableStartColumn;
                foreach (var cell in dataRow)
                {
                    worksheet.Cell(row, dataColumn++).Value = cell;
                }
                row++;
            }

            int lastRow = row - 1;
            int lastColumn = headerColumn - 1;

            worksheet.Range(startRow, tableStartColumn, lastRow, lastColumn).Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            worksheet.Range(startRow, tableStartColumn, lastRow, lastColumn).Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            worksheet.Range(startRow, tableStartColumn, lastRow, lastColumn).Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            worksheet.Range(startRow, tableStartColumn, lastRow, lastColumn).Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
        }
    }
}
