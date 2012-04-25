using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pickles.Parser;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelScenarioFormatter
    {
        private readonly ExcelStepFormatter excelStepFormatter;

        public ExcelScenarioFormatter(ExcelStepFormatter excelStepFormatter)
        {
            this.excelStepFormatter = excelStepFormatter;
        }

        public void Format(IXLWorksheet worksheet, Pickles.Parser.Scenario scenario, ref int row)
        {
            worksheet.Cell(row++, "B").Value = scenario.Name;
            worksheet.Cell(row++, "C").Value = scenario.Description;

            foreach (var step in scenario.Steps)
            {
                this.excelStepFormatter.Format(worksheet, step, ref row);
            }
        }
    }
}
