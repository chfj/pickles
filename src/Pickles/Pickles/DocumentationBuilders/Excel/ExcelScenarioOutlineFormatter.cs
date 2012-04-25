using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pickles.DocumentationBuilders.Excel
{
    public class ExcelScenarioOutlineFormatter
    {
        private readonly ExcelStepFormatter excelStepFormatter;
        private readonly ExcelTableFormatter excelTableFormatter;

        public ExcelScenarioOutlineFormatter(ExcelStepFormatter excelStepFormatter, ExcelTableFormatter excelTableFormatter)
        {
            this.excelStepFormatter = excelStepFormatter;
            this.excelTableFormatter = excelTableFormatter;
        }

        public void Format(ClosedXML.Excel.IXLWorksheet worksheet, Parser.ScenarioOutline scenarioOutline, ref int row)
        {
            worksheet.Cell(row++, "B").Value = scenarioOutline.Name;
            worksheet.Cell(row++, "C").Value = scenarioOutline.Description;

            foreach (var step in scenarioOutline.Steps)
            {
                this.excelStepFormatter.Format(worksheet, step, ref row);
            }

            row++;
            worksheet.Cell(row++, "B").Value = "Examples";
            this.excelTableFormatter.Format(worksheet, scenarioOutline.Example.TableArgument, ref row);
        }
    }
}
