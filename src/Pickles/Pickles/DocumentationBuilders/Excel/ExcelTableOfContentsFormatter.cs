using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelTableOfContentsFormatter
    {
        public void Format(XLWorkbook workbook)
        {
            var worksheet = workbook.AddWorksheet("TOC", 0);

            int row = 1;
            foreach (var featureWorksheet in workbook.Worksheets.Where(sheet => sheet.Name != "TOC"))
            {
                worksheet.Cell(row, 1).Value = featureWorksheet.Cell("A1").Value;
                worksheet.Cell(row++, 1).Hyperlink = new XLHyperlink(featureWorksheet.Name + "!A1");
            }
        }
    }
}
