using System;
using System.Collections.Generic;
using System.IO;

using FileDiffLib;

using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using OfficeOpenXml;

namespace IgxlLib.Test.UT.IgxlReader
{
    [TestClass]
    public class IgxlDataReaderTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void RemoveDuplicateInstanceTest()
        {
            string subName = "RemoveDuplicateInstanceTest";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var readInstanceSheet = new ReadInstanceSheet();
            string fileName = Path.Combine(InputPath, "TestInst_Common.txt");
            InstanceSheet instanceSheet = readInstanceSheet.GetSheet(fileName);

            instanceSheet.RemoveDuplicateInstance();
            string outputFile = Path.Combine(outputPath, "TestInst_Common.txt");
            instanceSheet.Write(outputFile, "3.0");

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void Stream_ReadIgxlSheetTest()
        {
            string subName = "ReadIgxlSheet";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var testCases = new Dictionary<string, Type>
            {
                { "AC_Specs.txt", typeof(ReadAcSpecSheet) },
                { "DC_Specs_CP1.txt", typeof(ReadDcSpecSheet) },
                { "Global Specs.txt", typeof(ReadGlobalSpecSheet) },
                { "PinMap.txt", typeof(ReadPinMapSheet) },
                { "PortMap_UF.txt", typeof(ReadPortMapSheet) },
                { "ChannelMap_4_site_CP.txt", typeof(ReadChanMapSheet) },
                { "Flow_nWire_Default.txt", typeof(ReadFlowSheet) },
                { "TestInst_Common.txt", typeof(ReadInstanceSheet) },
                { "Levels_nWire.txt", typeof(ReadLevelSheet) },
                { "TIMESET_nWire.txt", typeof(ReadTimeSetSheet) },
                { "JobList.txt", typeof(ReadJobListSheet) },
                { "References.txt", typeof(ReadReferenceSheet) },
                { "PatSets_All.txt", typeof(ReadPatSetSheet) },
                { "Pattern_Subroutine.txt", typeof(ReadPatSubroutineSheet) },
                { "Bin_Table.txt", typeof(ReadBinTableSheet) }
            };

            foreach (KeyValuePair<string, Type> testCase in testCases)
            {
                string fileName = testCase.Key;
                Type readerType = testCase.Value;
                string inputFilePath = Path.Combine(InputPath, fileName);
                object reader = Activator.CreateInstance(readerType);
                dynamic sheet = ((dynamic)reader).GetSheet(inputFilePath);
                string outputFile = Path.Combine(outputPath, fileName);
                sheet.Write(outputFile);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void ExcelSheet_ReadIgxlSheetTest()
        {
            string subName = "ReadIgxlSheet";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var testCases = new Dictionary<string, Type>
            {
                { "AC_Specs.txt", typeof(ReadAcSpecSheet) },
                { "DC_Specs_CP1.txt", typeof(ReadDcSpecSheet) },
                { "Global Specs.txt", typeof(ReadGlobalSpecSheet) },
                { "PinMap.txt", typeof(ReadPinMapSheet) },
                { "PortMap_UF.txt", typeof(ReadPortMapSheet) },
                { "ChannelMap_4_site_CP.txt", typeof(ReadChanMapSheet) },
                { "Flow_nWire_Default.txt", typeof(ReadFlowSheet) },
                { "TestInst_Common.txt", typeof(ReadInstanceSheet) },
                { "Levels_nWire.txt", typeof(ReadLevelSheet) },
                { "TIMESET_nWire.txt", typeof(ReadTimeSetSheet) },
                { "JobList.txt", typeof(ReadJobListSheet) },
                { "References.txt", typeof(ReadReferenceSheet) },
                { "PatSets_All.txt", typeof(ReadPatSetSheet) },
                { "Pattern_Subroutine.txt", typeof(ReadPatSubroutineSheet) },
                { "Bin_Table.txt", typeof(ReadBinTableSheet) }
            };

            foreach (KeyValuePair<string, Type> testCase in testCases)
            {
                string fileName = testCase.Key;
                Type readerType = testCase.Value;
                string inputFilePath = Path.Combine(InputPath, fileName);
                object reader = Activator.CreateInstance(readerType);
                ExcelWorksheet excelSheet = IgxlSheetReaderHelpers.ConvertTxtToExcelSheet(inputFilePath);
                dynamic sheet = ((dynamic)reader).ReadSheet(excelSheet);
                string outputFile = Path.Combine(outputPath, fileName);
                sheet.Write(outputFile);
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
