using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TechTalk.SpecFlow;
using TestAutomation.Framework.Helpers.Assertions;
using TestAutomation.Framework.Helpers.Excel;
using TestAutomation.Framework.Helpers.UI_Helpers;

namespace TestAutomation.ExcelTests
{
    [Binding]
    public sealed class TestExcelFeaturesSteps: UIFramework
    {
        private static string xlWorkbookPath = "TestExcel.xlsx";
        private readonly ExcelInteropHelper excelInterop;        

        public TestExcelFeaturesSteps()
        {
           excelInterop = new ExcelInteropHelper(xlWorkbookPath);
        }

        [Given(@"I have valid workbook")]
        public void GivenIHaveValidWorkbook()
        {
            ExcelInteropHelper.ActivateSheet();
        }

        [Then(@"I should get active workheet name")]
        public void ThenIShouldGetActiveWorkheetName()
        {
            string activeSheetName = ExcelInteropHelper.GetActiveSheetName();
            AssertHelpers.AssertEquals("Test1", activeSheetName);
        }


        [Then(@"I should get all workheet names")]
        public void ThenIShouldGetAllWorkheetNames()
        {
            List<string> actualSheetNames = ExcelInteropHelper.GetAllSheetNames();
            var expectedSheetNames = new List<string>();
            expectedSheetNames.AddRange(new[] { "Test1", "Test2", "IP (1) Process design or execut", "RiskRegisterWithScope" });

            CollectionAssert.AreEquivalent(expectedSheetNames, actualSheetNames);
        }

        [Then(@"I can activate worksheets by name/ index")]
        public void ThenICanActivateWorksheetsByNameIndex()
        {
            ExcelInteropHelper.ActivateSheet();
            Assert.AreEqual("Test1", ExcelInteropHelper.GetActiveSheetName());

            ExcelInteropHelper.ActivateSheet(2);
            Assert.AreEqual("Test2", ExcelInteropHelper.GetActiveSheetName());

            ExcelInteropHelper.ActivateSheet("Test1");
            Assert.AreEqual("Test1", ExcelInteropHelper.GetActiveSheetName());
        }

        [Then(@"I can read excel cell values of different format")]
        public void ThenICanReadExcelCellValuesOfDifferentFormat()
        {
            string cellValue = ExcelInteropHelper.ReadExcelCellText(1, 1);
            Assert.AreEqual("A", cellValue);

            // Using ColumnName
            cellValue = ExcelInteropHelper.ReadExcelCellText(1, "A");
            Assert.AreEqual("A", cellValue);

            // Using blank ColumnName
            cellValue = ExcelInteropHelper.ReadExcelCellText(1);
            Assert.AreEqual("A", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(2, 1);
            Assert.AreEqual("2.001", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(3, 1);
            Assert.AreEqual("5", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(4, 1);
            Assert.AreEqual("ABC", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(5, 1);
            Assert.AreEqual("XYZ", cellValue);

            // This is using formula
            //cellValue = ExcelInteropHelper.ReadExcelCellText(6, 1);
            //Assert.AreEqual("5", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(7, 1);
            Assert.AreEqual("₹ 23.75", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(8, 1);
            Assert.AreEqual("12/15/2019", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(9, 1);
            Assert.AreEqual("December 15, 2019", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(10, 1);
            Assert.AreEqual("8:52:48 AM", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(11, 1);
            Assert.AreEqual("76.90%", cellValue);

            cellValue = ExcelInteropHelper.ReadExcelCellText(12, 1);
            Assert.AreEqual("", cellValue);
        }

        [Then(@"I can get excel range co-ordinates")]
        public void ThenICanGetExcelRangeCo_Ordinates()
        {
            var actualValues = ExcelInteropHelper.GetRangeCoOrdinates("A1:A11");
            var expectedValues = new List<int>();
            expectedValues.AddRange(new[] { 1, 1, 1, 11 });

            CollectionAssert.AreEquivalent(expectedValues, actualValues);
        }

        [Then(@"I can get excel range as list")]
        public void ThenICanGetExcelRangeAsList()
        {
            var actualValues = ExcelInteropHelper.ReadExcelRangeProperty("A1:A11");

            var expectedValues = new List<string>();
            expectedValues.AddRange(new[] { "A", "2.001", "5", "ABC", "XYZ", "2.498750625", "₹ 23.75", "12/15/2019", "December 15, 2019", "8:52:48 AM", "76.90%" });
            CollectionAssert.AreEquivalent(expectedValues, actualValues);
        }

        [Then(@"I can get excel cell background color")]
        public void ThenICanGetExcelCellBackgroundColor()
        {
            var actualValues = ExcelInteropHelper.ReadExcelRangeProperty("A3", ExcelCellProperties.CellBackGroundColor);

            var expectedValues = new List<string>();
            expectedValues.AddRange(new[] { "ffffff" });
            CollectionAssert.AreEquivalent(expectedValues, actualValues);
        }

        [Then(@"I can get all excel cell properties")]
        public void ThenICanAllExcelCellProperties()
        {
            ExcelInteropHelper.ActivateSheet("RiskRegisterWithScope");
            ExcelInteropHelper.GetCellProperties(5, 8);
            string bgColor = ExcelCellProperties.BackGroundColor;
            string fontColor = ExcelCellProperties.FontColor;
            string fontSize = ExcelCellProperties.FontSize;
            string formula = ExcelCellProperties.Formula;
            string text = ExcelCellProperties.Text;

            ExcelInteropHelper.GetCellProperties("H5");
            bgColor = ExcelCellProperties.BackGroundColor;
        }

        [Then(@"I can get all excel range properties")]
        public void ThenICanGetAllExcelRangeProperties()
        {
            //ExcelInteropHelper.GetCellProperties(5, 1);
            ExcelInteropHelper.ActivateSheet("RiskRegisterWithScope");
            var bgColorProperty = ExcelInteropHelper.ReadExcelRangeProperty("H5:S5", ExcelCellProperties.CellBackGroundColor);
            var fontColorProperty = ExcelInteropHelper.ReadExcelRangeProperty("H5:S5", ExcelCellProperties.CellFontColor);
            var fontSizeProperty = ExcelInteropHelper.ReadExcelRangeProperty("H5:S5", ExcelCellProperties.CellFontSize);
            var formulaProperty = ExcelInteropHelper.ReadExcelRangeProperty("H5:S5", ExcelCellProperties.CellFormula);
            var textProperty = ExcelInteropHelper.ReadExcelRangeProperty("H5:S5", ExcelCellProperties.CellText);

            int index = 0;
            foreach (var text in textProperty)
            {
                if (Convert.ToInt32(text) < 2)
                {
                    Assert.AreEqual("ffffff", bgColorProperty[index]);
                }
                else if (Convert.ToInt32(text) >= 2 && Convert.ToInt32(text) < 5)
                {
                    Assert.AreEqual("ffffff", bgColorProperty[index]);
                }
                else if (Convert.ToInt32(text) >= 5)
                {
                    Assert.AreEqual("a31e22", bgColorProperty[index]);
                }

                index++;
            }
        }

        [Then(@"I can get all excel cell formula")]
        public void ThenICanGetAllExcelCellFormula()
        {
            string actualFormula = ExcelInteropHelper.ReadExcelFormula(6, 1);
            Assert.AreEqual("=A3/A2", actualFormula);
        }

        [Then(@"I can write to excel")]
        public void ThenICanWriteToExcel()
        {
            ExcelInteropHelper.WriteExcel(1, 9, "test");
            //ExcelInteropHelper.SaveExcel(); 

            // 1. Single values
            ExcelInteropHelper.WriteExcel("I2", "test1");
            //ExcelInteropHelper.CloseExcel();

            // 2. Multidimentional Arrays
            string[,] rangeValues = new string[3, 3] {{"a","b","c"},
                                      {"d","e","f"},
                                      {"g","h","i"} };

            ExcelInteropHelper.WriteExcel("I3:K5", rangeValues);
            //ExcelInteropHelper.CloseExcel();

            // When values are of dynamic range, mention start cell only
            ExcelInteropHelper.WriteExcel("I6", rangeValues);

            rangeValues = new string[3, 1] {  {"a"},
                                              {"d"},
                                              {"g"} };
            // When values are of dynamic range, mention start cell only
            ExcelInteropHelper.WriteExcel("I9", rangeValues);

            // 3. One dimensional array (single column)
            string[] rangeValues1 = new string[3] { "a", "b", "c" };
            ExcelInteropHelper.WriteExcel("I12", rangeValues1);

            // Always use CloseExcel at end of all oprations
            ExcelInteropHelper.CloseExcel();
        }

        [Then(@"I can read chart values")]
        public void ThenICanReadChartValues()
        {
            ExcelInteropHelper.ActivateSheet("IP (1) Process design or execut");
            ExcelInteropHelper.ReadExcelChart();
        }

        [Then(@"I can get all excel cell font color")]
        public void ThenICanGetAllExcelCellFontColor()
        {
            ExcelInteropHelper.SetExcelRangeProperty("A1:A15", ExcelCellProperties.CellFontColor, "#FF0000");
            ExcelInteropHelper.CloseExcel();
        }

        [Then(@"I can get all excel cell bold property")]
        public void ThenICanGetAllExcelCellBoldProperty()
        {
            ExcelInteropHelper.SetExcelRangeProperty("A1:A15", ExcelCellProperties.CellFontBold, "Bold");
            ExcelInteropHelper.CloseExcel();
        }

        [Then(@"I can set excel cell border")]
        public void ThenICanSetExcelCellBorder()
        {
            ExcelInteropHelper.SetBorderAround("A1:A15");
            ExcelInteropHelper.CloseExcel();
        }


    }
}
