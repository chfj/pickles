using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelSheetNameGenerator
    {
        public string GenerateSheetName(ClosedXML.Excel.XLWorkbook workbook, Parser.Feature feature)
        {
            string name = feature.Name.Replace(" ", string.Empty).Replace("\t", string.Empty).ToUpperInvariant();
            if (name.Length > 31) name = name.Substring(0, 31);

            // check if the workbook contains any sheets with this name
            int nextIndex = 1;
            while (workbook.Worksheets.Any(sheet => sheet.Name == name))
            {
                name = name.Remove(name.Length - 3, 3);
                name = name + "(" + nextIndex++ + ")";
            }

            return name;
        }
    }
}
