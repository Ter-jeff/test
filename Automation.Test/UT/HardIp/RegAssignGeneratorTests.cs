using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class RegAssignGeneratorTests : FunctionTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            LocalSpecs.TarFolder = OutputPath;
        }

        private static List<HardIpRegAssign> CreateFakeRegAssignList(RegisterAssignType registerAssignType = RegisterAssignType.ForceV_Val)
        {
            var item1 = new HardIpRegAssign
            {
                Type = registerAssignType,
                SubBlockName = "Item1_Sub1",
                Headers = ["RegName", "Value"],
                RegAssignList =
                [
                    ["REG_1", "0x01"],
                    ["REG_2", "0x02"]
                ]
            };

            var item2 = new HardIpRegAssign
            {
                Type = registerAssignType,
                SubBlockName = "Item2_Sub1",
                Headers = ["RegName", "Value"],
                RegAssignList =
                [
                    ["REG_1", "0x01"],
                    ["REG_2", "0x02"]
                ]
            };
            return [item1, item2];
        }

        [TestMethod]
        public void WorkFlowTest()
        {
            string subName = "RegAssignGeneratorWorkFlow";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            List<HardIpRegAssign> assigns = CreateFakeRegAssignList();
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpInputData = new HardIpInputData(hardIpParaData);
            var gen = new RegAssignGenerator(hardIpInputData);
            using (var package = new ExcelPackage())
            {
                // Act
                gen.WorkFlow(package.Workbook, assigns);

                // Assert
                ExcelWorksheet ws = package.Workbook.Worksheets["Reg_Assign_Item1"];
                ws.ExportToTxt(Path.Combine(outputPath, ws.Name + ".txt"));

                ws = package.Workbook.Worksheets["Reg_Assign_Item2"];
                ws.ExportToTxt(Path.Combine(outputPath, ws.Name + ".txt"));
            }
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void WorkFlow_EfuseReadWriteType_GroupsUnderEfuseSheetAndUsesInstanceNameLabel()
        {
            // Arrange - HIP_eFuse_Read/Write items are grouped under "_eFuse" regardless of
            // SubBlockName, and labeled "InstanceName" instead of "SubBlockName"
            var item = new HardIpRegAssign
            {
                Type = RegisterAssignType.HIP_eFuse_Read,
                SubBlockName = "AnySub_X",
                RegAssignList = [["Cat1", "Code1"]]
            };
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpInputData = new HardIpInputData(hardIpParaData);
            var gen = new RegAssignGenerator(hardIpInputData);

            using var package = new ExcelPackage();

            // Act
            gen.WorkFlow(package.Workbook, [item]);

            // Assert
            ExcelWorksheet ws = package.Workbook.Worksheets["Reg_Assign_eFuse"];
            Assert.IsNotNull(ws);
            Assert.AreEqual("InstanceName", ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void WorkFlow_NonEfuseType_UsesSubBlockNameLabel()
        {
            // Arrange
            var item = new HardIpRegAssign
            {
                Type = RegisterAssignType.ForceV_Val,
                SubBlockName = "Item1_Sub1",
                RegAssignList = [["REG_1", "0x01"]]
            };
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpInputData = new HardIpInputData(hardIpParaData);
            var gen = new RegAssignGenerator(hardIpInputData);

            using var package = new ExcelPackage();

            // Act
            gen.WorkFlow(package.Workbook, [item]);

            // Assert
            ExcelWorksheet ws = package.Workbook.Worksheets["Reg_Assign_Item1"];
            Assert.IsNotNull(ws);
            Assert.AreEqual("SubBlockName", ws.Cells[2, 1].Value);
        }

        [TestMethod]
        public void WorkFlow_HeadersNullType_FallsBackToRegAssignListColumnCount()
        {
            // Arrange - NoneType is not handled by the Headers switch, so Headers stays null
            // and WriteData must fall back to RegAssignList's own max column count
            var item = new HardIpRegAssign
            {
                Type = RegisterAssignType.NoneType,
                SubBlockName = "Item1_Sub1",
                RegAssignList = [["A", "B", "C"]]
            };
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpInputData = new HardIpInputData(hardIpParaData);
            var gen = new RegAssignGenerator(hardIpInputData);

            using var package = new ExcelPackage();

            // Act
            gen.WorkFlow(package.Workbook, [item]);

            // Assert
            ExcelWorksheet ws = package.Workbook.Worksheets["Reg_Assign_Item1"];
            Assert.IsNotNull(ws);
            Assert.AreEqual("A", ws.Cells[3, 2].Value);
            Assert.AreEqual("B", ws.Cells[3, 3].Value);
            Assert.AreEqual("C", ws.Cells[3, 4].Value);
        }
    }
}
