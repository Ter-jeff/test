using System;
using System.Collections.Generic;
using System.IO;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenConti.Base;
using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;
using Automation.Reader;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.Basic
{
    [TestClass]
    public class ContiMainTests : FunctionTestBase
    {
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void ContiMainTest()
        {
            string subName = "ContiMain";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);
            LocalSpecs.TarFolder = outputPath;

            // Arrange
            var testCases = new List<(string Category, ContiType Expected)>
            {
                ("continuity check", ContiType.OpenShort),
                ("power short", ContiType.PowerShort),
                ("power sense", ContiType.PowerSense),
                ("power impedance", ContiType.PowerImpedance),
                ("ground sense", ContiType.GroundSense),
                ("ground impedance", ContiType.GroundImpedance),
                ("continuity analog", ContiType.ContiAnalog),
                ("walking z", ContiType.WalkingZ),
                ("userfunction", ContiType.UserFunction),
                ("ContinuityPPMU", ContiType.ContiPpmu),
                ("Opcode",  ContiType.Opcode),
                ("AutoZTtr",  ContiType.AutoZTtr)
            };

            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            // Act
            var dcTestContiRows = new List<DcTestContiRow>();
            foreach ((string category, ContiType _) in testCases)
            {
                dcTestContiRows.Add(new DcTestContiRow
                {
                    Category = category,
                    JobNameList = ["CP1", "CP2"],
                    PinGroup = "All_IO",
                    Limits =
                        [
                            new()
                            {
                                LimitStage = "1",
                                LimitHeader = "1",
                                LimitValue = "1",
                                LimitValueSecondary = "1",
                                LimitType = "1",
                                HiLimitValue = "1",
                                LoLimitValue = "1",
                                ForceConditionValue = "1"
                            }
                        ]
                });
            }
            dcTestContiSheet.DcTestContiRows = dcTestContiRows;
            var contiMain = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);
            List<BinTableRow> binTableRows = [];
            ContiResult results = contiMain.WorkFlow(null, ref binTableRows);
            results.ContiBinTableRows = binTableRows;

            Print(Path.Combine(OutputPath, "Basic", subName), results);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        public static void Print(string outputPath, ContiResult contiResult)
        {
            contiResult.ContiFlowSheet.Write(Path.Combine(outputPath, contiResult.ContiFlowSheet.Name + ".txt"));
            contiResult.ConInstanceSheet.Write(Path.Combine(outputPath, contiResult.ConInstanceSheet.Name + ".txt"));
            contiResult.CommonInstance.Write(Path.Combine(outputPath, contiResult.CommonInstance.Name + ".txt"));
            var binTable = new BinTableSheet("BinTable");
            binTable.AddRows(contiResult.ContiBinTableRows);
            binTable.Write(Path.Combine(outputPath, binTable.Name + ".txt"));
        }

        [TestMethod]
        public void DcTestContiRowTypeTest()
        {
            // Arrange
            var testCases = new List<(string Category, ContiType Expected)>
            {
                ("continuity check", ContiType.OpenShort),
                ("power short", ContiType.PowerShort),
                ("power sense", ContiType.PowerSense),
                ("power impedance", ContiType.PowerImpedance),
                ("ground sense", ContiType.GroundSense),
                ("ground impedance", ContiType.GroundImpedance),
                ("continuity analog", ContiType.ContiAnalog),
                ("power open resistance", ContiType.Cres),
                ("walking z", ContiType.WalkingZ),
                ("userfunction", ContiType.UserFunction),
                ("unknown category", ContiType.UnKnow),
                ("ContinuityPPMU", ContiType.ContiPpmu),
                ("Opcode",  ContiType.Opcode),
                ("AutoZTtr",  ContiType.AutoZTtr)
            };

            // Assert
            foreach ((string category, ContiType expected) in testCases)
            {
                var row = new DcTestContiRow { Category = category };
                ContiType result = row.TestType;
                Assert.AreEqual(expected, result, $"Category '{category}' should map to {expected}, but got {result}.");
            }
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void DcUseFunctionTest()
        {
            string subName = "ContiMain";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            LocalSpecs.TarFolder = outputPath;
            var testCases = new List<(string Category, ContiType Expected)>
            {
                ("userfunction", ContiType.UserFunction),
                ("power short", ContiType.PowerShort)
            };

            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            // Act
            var dcTestContiRows = new List<DcTestContiRow>();
            foreach ((string category, ContiType _) in testCases)
            {
                dcTestContiRows.Add(new DcTestContiRow
                {
                    Category = category,
                    JobNameList = ["CP1", "CP2"],
                    PinGroup = "All_IO",
                    Condition = "opcode=test",
                    FailFlag = "F_userfunction",

                    Limits =
                        [
                            new()
                            {
                                LimitStage = "1",
                                LimitHeader = "1",
                                LimitValue = "1",
                                LimitValueSecondary = "1",
                                LimitType = "1",
                                HiLimitValue = "1",
                                LoLimitValue = "1",
                                ForceConditionValue = "1"
                            }
                        ]
                });
            }

            dcTestContiSheet.DcTestContiRows = dcTestContiRows;
            var contiMain = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);
            List<BinTableRow> binTableRows = [];
            ContiResult results = contiMain.WorkFlow(null, ref binTableRows);
            results.ContiBinTableRows = binTableRows;

            Print(Path.Combine(OutputPath, "Basic", subName), results);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [DataRow(false)]
        [DataRow(true)]
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void DcPowerShortNoPinGrpTest(bool isParallel)
        {
            string subName = "ContiMain";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            LocalSpecs.TarFolder = outputPath;
            string extraString = isParallel ? "parallel" : "serial";
            var testCases = new List<(string Category, ContiType Expected)>
            {
                ($"power short {extraString}", ContiType.PowerShort)
            };

            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            // Act
            var dcTestContiRows = new List<DcTestContiRow>();
            foreach ((string category, ContiType _) in testCases)
            {
                dcTestContiRows.Add(new DcTestContiRow
                {
                    Category = category,
                    JobNameList = ["CP1", "CP2"],
                    PinGroup = "All_IO",
                    EnableWord = "PowerShort",

                    Limits =
                        [
                            new()
                            {
                                LimitStage = "1",
                                LimitHeader = "1",
                                LimitValue = "1",
                                LimitValueSecondary = "1",
                                LimitType = "1",
                                HiLimitValue = "1",
                                LoLimitValue = "1",
                                ForceConditionValue = "1"
                            }
                        ]
                });
            }

            if (isParallel)
            {
                Function function = new Function { FunctionName = DcContiConst.CSharpFuncNamePowerShort, Type = ".NET" };
                TestProgram.VbtFunctionLib.AddVbtFunction(function);
            }
            else
            {
                TestProgram.VbtFunctionLib.VbtLib.Clear();
            }

            dcTestContiSheet.DcTestContiRows = dcTestContiRows;
            var contiMain = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);
            List<BinTableRow> binTableRows = [];
            ContiResult results = contiMain.WorkFlow(null, ref binTableRows);
            results.ContiBinTableRows = binTableRows;

            Print(Path.Combine(OutputPath, "Basic", subName), results);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [DataRow(false)]
        [DataRow(true)]
        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void DcPowerSenseTest(bool isCSharp)
        {
            string subName = "ContiMain";
            string outputPath = Path.Combine(OutputPath, "Evs", subName);
            string expectPath = Path.Combine(ExpectPath, "Evs", subName);

            LocalSpecs.TarFolder = outputPath;
            var testCases = new List<(string Category, ContiType Expected)>
            {
                ("power sense", ContiType.PowerSense),
                ("power short", ContiType.PowerShort)
            };

            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            // Act
            var dcTestContiRows = new List<DcTestContiRow>();
            foreach ((string category, ContiType _) in testCases)
            {
                dcTestContiRows.Add(new DcTestContiRow
                {
                    Category = category,
                    JobNameList = ["CP1", "CP2"],
                    EnableWord = "PowerSense",
                    Condition = "Iforce=10m",
                    PinGroup = "All_IO",

                    Limits =
                    [
                        new()
                        {
                            LimitStage = "1",
                            LimitHeader = "1",
                            LimitValue = "1",
                            LimitValueSecondary = "1",
                            LimitType = "1",
                            HiLimitValue = "1",
                            LoLimitValue = "1",
                            ForceConditionValue = "1"
                        }
                    ]
                });
            }

            if (isCSharp)
            {
                Function function = new Function
                {
                    FunctionName = DcContiConst.CSharpFuncNameSensePinConti,
                    Type = ".NET"
                };
                TestProgram.VbtFunctionLib.AddVbtFunction(function);
            }
            else
            {
                TestProgram.VbtFunctionLib.VbtLib.Clear();
            }

            dcTestContiSheet.DcTestContiRows = dcTestContiRows;
            var contiMain = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);
            List<BinTableRow> binTableRows = [];
            ContiResult results = contiMain.WorkFlow(null, ref binTableRows);
            results.ContiBinTableRows = binTableRows;

            Print(Path.Combine(OutputPath, "Basic", subName), results);

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void RemoveBinOutFlowRowsIfNeeded_Test()
        {
            // Arrange
            var flowRows = new List<FlowRow>
            {
                null!,
                new() { Opcode = "Bin", Parameter = "x" },
                null!,
                new() { Opcode = "Test", Parameter = "y" }
            };

            // Act
            int removedCount = ContiMain.RemoveBinOutFlowRowsIfNeeded(flowRows, "true");

            // Assert
            Assert.AreEqual(1, removedCount);
            Assert.AreEqual(3, flowRows.Count);
            Assert.AreEqual("Test", flowRows[2].Opcode);
        }

        [TestMethod]
        public void RemoveBinOutFlowRowsIfNeeded_FlowRowsIsNull_ShouldNotThrow()
        {
            // Arrange
            List<FlowRow>? flowRows = null;

            // Act
            ContiMain.RemoveBinOutFlowRowsIfNeeded(flowRows, "true");

            // Assert
            // No exception thrown
        }

        [TestMethod]
        public void RemoveBinOutFlowRowsIfNeeded_NullItem_ShouldNotThrow()
        {
            // Arrange
            var flowRows = new List<FlowRow>
            {
                null!
            };

            // Act
            int removedCount = ContiMain.RemoveBinOutFlowRowsIfNeeded(flowRows, "true");

            // Assert
            Assert.AreEqual(0, removedCount);
            Assert.AreEqual(1, flowRows.Count);
        }

        [TestMethod]
        public void AllPPMUFlag_NoContinuity_AnyShouldBeFalse()
        {
            // Arrange
            List<ContiBase> contiTestList = BuildContiBases(includeContinuity: false);

            // Act
            List<string> allPPMU_Flag = ContiMain.GetAllContinuityFlags(contiTestList);

            // Assert
            Assert.IsFalse(allPPMU_Flag.Count != 0);
        }

        [TestMethod]
        public void AllPPMUFlag_WithContinuity_AnyShouldBeTrue()
        {
            // Arrange
            List<ContiBase> contiTestList = BuildContiBases(includeContinuity: true);

            // Act
            List<string> allPPMU_Flag = ContiMain.GetAllContinuityFlags(contiTestList);

            // Assert
            Assert.IsTrue(allPPMU_Flag.Count != 0);
        }

        private static List<ContiBase> BuildContiBases(bool includeContinuity)
        {
            var list = new List<ContiBase>();

            var row1 = new DcTestContiRow { Category = "IDD", SiteFlag = "S1" };
            list.Add(new ContiPowerShort(row1));

            var row2 = new DcTestContiRow { Category = "Leakage", SiteFlag = "S2" };
            list.Add(new ContiPowerShort(row2));

            if (includeContinuity)
            {
                var row3 = new DcTestContiRow
                {
                    Category = "Continuity",
                    SiteFlag = "S3"
                };
                list.Add(new ContiPowerShort(row3));
            }

            return list;
        }

        [TestMethod]
        public void BuildFlagString_NoFlags_ShouldReturnNull()
        {
            // Arrange
            var flags = new List<string>();

            // Act
            string result = ContiMain.BuildFlagString(flags);

            // Assert
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void BuildFlagString_OneFlag_ShouldReturnThatFlag()
        {
            var flags = new List<string> { "S1" };

            string result = ContiMain.BuildFlagString(flags);

            Assert.AreEqual("S1", result);
        }

        [TestMethod]
        public void BuildFlagString_MultipleFlags_ShouldReturnWrappedString()
        {
            // Arrange
            var flags = new List<string> { "S1", "S2" };

            // Act
            string result = ContiMain.BuildFlagString(flags);

            // Assert
            Assert.AreEqual("S1, S2", result);
        }

        [TestMethod]
        public void HandlePPMUFlag_WithContinuity_ShouldAddRow()
        {
            // Arrange
            List<ContiBase> contiTestList = BuildContiBases(includeContinuity: true);

            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Flow");
            var contiFlow = new SubFlowSheet(ws);

            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);
            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            // Act
            sut.HandlePPMUFlag(contiTestList, contiFlow);

            // Assert
            Assert.AreEqual(1, contiFlow.Rows.Count);
        }

        [TestMethod]
        public void HandlePPMUFlag_NoContinuity_ShouldNotAddFlowRow()
        {
            // Arrange
            List<ContiBase> contiTestList = BuildContiBases(includeContinuity: false);

            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Flow");
            var contiFlow = new SubFlowSheet(ws);

            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);
            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            // Act
            sut.HandlePPMUFlag(contiTestList, contiFlow);

            // Assert
            Assert.AreEqual(0, contiFlow.Rows.Count);
        }

        [TestMethod]
        public void AddMapping_SameValueMultiplePins_ShouldGroupCorrectly()
        {
            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);
            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            var conditions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "AAA", "A,B,C" },
                { "BBB", "1.2,1.2,1.3" }
            };

            var result = new Dictionary<double, List<string>>();

            sut.ConfigureClampLimits(conditions, result, "AAA", "BBB");

            Assert.AreEqual(2, result.Count);

            CollectionAssert.AreEqual(
                new List<string> { "A", "B" },
                result[1.2]);

            CollectionAssert.AreEqual(
                new List<string> { "C" },
                result[1.3]);
        }

        [TestMethod]
        public void AddMapping_KeyNotFound_ShouldNotThrow()
        {
            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);
            var conditions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var result = new Dictionary<double, List<string>>();

            sut.ConfigureClampLimits(conditions, result, "AAA", "BBB");

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GetClampSettingFromContiPlanByUser_ShouldParseClampSettingSuccessfully()
        {
            // Arrange
            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            List<ContiBase> contiTestList = [];
            var row1 = new DcTestContiRow
            {
                InstanceName = "SetPPMU_Clamp_Conti",
                Condition = "clampHiGroup=A,B,C;vch=1.2,1.2,1.3;clampLoGroup=D,E;vcl=-1.2,-1.3"
            };
            contiTestList.Add(new ContiPowerShort(row1));

            var clampLimit = new Dictionary<double, List<string>>();

            // Act
            sut.GetClampSettingFromContiPlanByUser(contiTestList, ref clampLimit);

            // Assert
            Assert.AreEqual(4, clampLimit.Count);

            CollectionAssert.AreEquivalent(
                new List<string> { "A", "B" },
                clampLimit[1.2]);

            CollectionAssert.AreEquivalent(
                new List<string> { "C" },
                clampLimit[1.3]);

            CollectionAssert.AreEquivalent(
                new List<string> { "D" },
                clampLimit[-1.2]);

            CollectionAssert.AreEquivalent(
                new List<string> { "E" },
                clampLimit[-1.3]);
        }

        [TestMethod]
        public void GetClampSettingFromContiPlanByUser_InstanceNotFound_ShouldNotThrow()
        {
            // Arrange
            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            List<ContiBase> contiTestList = [];
            var row1 = new DcTestContiRow
            {
                InstanceName = "OtherInstance",
                Condition = "AAA=BBB"
            };
            contiTestList.Add(new ContiPowerShort(row1));

            var clampLimit = new Dictionary<double, List<string>>();

            // Act
            sut.GetClampSettingFromContiPlanByUser(contiTestList, ref clampLimit);

            // Assert
            Assert.AreEqual(0, clampLimit.Count);
        }

        [TestMethod]
        public void GetClampSettingFromContiPlanByUser_NullCondition_ShouldNotThrow()
        {
            // Arrange
            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            List<ContiBase> contiTestList = [];
            var row1 = new DcTestContiRow
            {
                InstanceName = "SetPPMU_Clamp_Conti",
                Condition = string.Empty
            };
            contiTestList.Add(new ContiPowerShort(row1));

            var clampLimit = new Dictionary<double, List<string>>();

            // Act
            sut.GetClampSettingFromContiPlanByUser(contiTestList, ref clampLimit);

            // Assert
            Assert.AreEqual(0, clampLimit.Count);

        }

        [TestMethod]
        public void GetClampSettingFromContiPlanByUser_InstanceNotFound_ShouldNotAddClampLimit()
        {
            // Arrange
            DcTestContiSheet dcTestContiSheet = new DcTestContiSheet("");
            PatSetSheet patSetAll = new PatSetSheet("");
            IoInfoSheet ioInfoSheet = new IoInfoSheet("", []);

            var sut = new ContiMain(dcTestContiSheet, patSetAll, ioInfoSheet);

            List<ContiBase> contiTestList = [];
            var row1 = new DcTestContiRow
            {
                InstanceName = "OtherInstance",
                Condition = "clampHiGroup=A;vch=1.2"
            };
            contiTestList.Add(new ContiPowerShort(row1));

            var clampLimit = new Dictionary<double, List<string>>();

            // Act
            sut.GetClampSettingFromContiPlanByUser(contiTestList, ref clampLimit);

            // Assert
            Assert.AreEqual(0, clampLimit.Count);
        }

        [TestCleanup]
        public void ResetJobInfoSheet()
        {
            SetJobInfoSheet(null);
        }

        private static void SetJobInfoSheet(JobInfoSheet? sheet)
        {
            TestPlanStatic._jobInfoSheet = sheet;
        }

        [TestMethod]
        public void GetTesterTypeOrDefault_EmptyJobInfoSheet_ReturnsDefaultType()
        {
            // Arrange - a non-null but empty sheet is used (rather than null) so the cached
            // value short-circuits TestPlanStatic.JobInfoSheet's lazy real-workbook read,
            // which is otherwise order-dependent on other tests sharing the static workbook.
            SetJobInfoSheet(new JobInfoSheet());

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("CP1");

            // Assert
            Assert.AreEqual("UF", result);
        }

        [TestMethod]
        public void GetTesterTypeOrDefault_JobFoundWithTesterType_ReturnsUppercasedTesterType()
        {
            // Arrange
            var sheet = new JobInfoSheet();
            sheet.Rows.Add(new JobInfoRow { JobName = "CP1", TesterType = "uflex" });
            SetJobInfoSheet(sheet);

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("CP1");

            // Assert
            Assert.AreEqual("UFLEX", result);
        }

        [TestMethod]
        public void GetTesterTypeOrDefault_JobNotFound_ReturnsDefaultType()
        {
            // Arrange
            var sheet = new JobInfoSheet();
            sheet.Rows.Add(new JobInfoRow { JobName = "FT1", TesterType = "uflex" });
            SetJobInfoSheet(sheet);

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("CP1");

            // Assert
            Assert.AreEqual("UF", result);
        }

        [TestMethod]
        public void GetTesterTypeOrDefault_JobFoundWithBlankTesterType_ReturnsDefaultType()
        {
            // Arrange
            var sheet = new JobInfoSheet();
            sheet.Rows.Add(new JobInfoRow { JobName = "CP1", TesterType = "  " });
            SetJobInfoSheet(sheet);

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("CP1");

            // Assert
            Assert.AreEqual("UF", result);
        }

        [TestMethod]
        public void GetTesterTypeOrDefault_CustomDefaultType_UsedWhenJobNotFound()
        {
            // Arrange
            SetJobInfoSheet(new JobInfoSheet());

            // Act
            string result = ContiMain.GetTesterTypeOrDefault("CP1", "CUSTOM");

            // Assert
            Assert.AreEqual("CUSTOM", result);
        }
    }
}
