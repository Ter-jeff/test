using System;
using System.Collections.Generic;
using System.IO;

using Automation.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Utility
{
    [TestClass]
    public class MergeSheetTest : FunctionTestBase
    {
        private ExcelWorkbook? _workbook = null;
        public static readonly string VoltageTablePath = Path.Combine(Directory.GetCurrentDirectory(), "Input", "Utility", "MergeSheet");

        [TestInitialize]
        public void Setup()
        {
            _workbook = new ExcelPackage(new FileInfo(VoltageTablePath)).Workbook;
        }

        [TestMethod]
        public void MultipleVoltageTable()
        {
            bool isExcpetion = false;
            _workbook!.Worksheets.Add("Testsetting_CP1");
            List<string> voltageTables = [Path.Combine(VoltageTablePath, "Borneo_B0_VolTa_V07A_CP1_X#1.csv"), Path.Combine(VoltageTablePath, "Borneo_B0_VolTa_V07A_FT1_X#1.csv"), Path.Combine(VoltageTablePath, "Borneo_B0_VolTa_V07A_FT1_X#2.xlsx")];
            try
            {
                MergeSheet.ParseTestSettingSheetToTestPlan(voltageTables, _workbook);
            }
            catch (Exception)
            {
                isExcpetion = true;
            }
            Assert.IsTrue(isExcpetion);
        }

        [TestMethod]
        public void NonCsvVoltageTable()
        {
            _workbook!.Worksheets.Add("Testsetting_CP1");
            List<string> voltageTables = [Path.Combine(VoltageTablePath, "Borneo_B0_VolTa_V07A_CP1_X#2.xlsm"), Path.Combine(VoltageTablePath, "Borneo_B0_VolTa_V07A_FT1_X#2.xlsx")];
            MergeSheet.ParseTestSettingSheetToTestPlan(voltageTables, _workbook);
            Assert.IsTrue(_workbook.Worksheets.Count == voltageTables.Count);
        }
    }
}
