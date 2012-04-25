using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelStepFormatter
    {
        private readonly ExcelTableFormatter excelTableFormatter;
        private readonly ExcelDocumentStringFormatter excelDocumentStringFormatter;

        public ExcelStepFormatter(ExcelTableFormatter excelTableFormatter, ExcelDocumentStringFormatter excelDocumentStringFormatter)
        {
            this.excelTableFormatter = excelTableFormatter;
            this.excelDocumentStringFormatter = excelDocumentStringFormatter;
        }

        public void Format(ClosedXML.Excel.IXLWorksheet worksheet, Parser.Step step, ref int row)
        {
            worksheet.Cell(row++, "C").Value = step.NativeKeyword + " " + step.Name;

            if (step.TableArgument != null)
            {
                this.excelTableFormatter.Format(worksheet, step.TableArgument, ref row);
            }

            if (!string.IsNullOrEmpty(step.DocStringArgument))
            {
                this.excelDocumentStringFormatter.Format(worksheet, step.DocStringArgument, ref row);
            }
        }
    }
}
