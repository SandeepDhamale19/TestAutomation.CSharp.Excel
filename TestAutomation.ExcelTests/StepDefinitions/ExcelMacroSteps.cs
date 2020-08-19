using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TechTalk.SpecFlow;
using TestAutomation.Framework.Helpers.Excel;

namespace TestAutomation.ExcelTests.StepDefinitions
{
    [Binding]
    public sealed class ExcelMacroSteps
    {
        private static string xlWorkbookPath = "Report_Xxxxx_Corp_07Jan2020.xlsm";
        ExcelInteropHelper excelInterop = new ExcelInteropHelper(xlWorkbookPath, "GRS Before & After");

        string macroName = "Offline_Xxxxx_Corp_16Dec2019.xlsm!UserAgreement";

        public ExcelMacroSteps()
        {
            excelInterop = new ExcelInteropHelper(xlWorkbookPath);
        }

        [Given(@"I have macro enabled workbook")]
        public void GivenIHaveMacroEnabledWorkbook()
        {
            ExcelInteropHelper.ActivateSheet();
        }

        [Then(@"I can run macro")]
        public void ThenICanRunMacro()
        {
            ExcelInteropHelper.RunExcelMacro(new Object[] { macroName });
        }

    }
}
