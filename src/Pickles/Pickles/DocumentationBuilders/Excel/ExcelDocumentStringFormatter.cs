using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelDocumentStringFormatter
    {
        public void Format(ClosedXML.Excel.IXLWorksheet worksheet, string documentString, ref int row)
        {
            var documentStringLines = documentString.Split(new string[] { "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in documentStringLines)
            {
                worksheet.Cell(row++, 4).Value = line;
            }
        }
    }
}
