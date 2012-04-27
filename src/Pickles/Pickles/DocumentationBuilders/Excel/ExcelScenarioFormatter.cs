﻿#region License

/*
    Copyright [2011] [Jeffrey Cameron]

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

#endregion

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
            worksheet.Cell(row, "B").Style.Font.SetBold();
            worksheet.Cell(row++, "B").Value = scenario.Name;
            worksheet.Cell(row++, "C").Value = scenario.Description;

            foreach (var step in scenario.Steps)
            {
                this.excelStepFormatter.Format(worksheet, step, ref row);
            }
        }
    }
}
