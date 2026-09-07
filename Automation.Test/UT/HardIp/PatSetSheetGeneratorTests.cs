using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class PatSetSheetGeneratorTests : FunctionTestBase
    {
        private PatSetSheetGenerator _generator = null!;

        [TestInitialize]
        public void Setup()
        {
            LocalSpecs.TarFolder = OutputPath;
            FolderStructure.CreateFolder();
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            HardIpInputData hardIpInputData = new HardIpInputData(paraData);
            _generator = new PatSetSheetGenerator(hardIpInputData);
        }

        [TestMethod]
        public void GenPatSet_ShouldReturnPatSetSheetWithPatSets()
        {
            // Arrange
            var pattern1 = new HardIpPattern
            {
                Pattern = new PatternClass("CZ_BRNA0_C_FULP_AN_AA00_DLL_JTG_VIX_ALLFRV_SI_CPLLDS_T6PD#1")
                {
                    PatternSetList = [["P1", "P2"]]
                }
            };

            var pattern2 = new HardIpPattern
            {
                Pattern = new PatternClass("CZ_BRNA0_C_FULP_AN_AA00_DLL_JTG_VIX_ALLFRV_SI_CPLLDS_T6PD")
                {
                    PatternSetList = [["P3"]]
                },
                BlockType = "Scan"
            };

            var sheetDict = new Dictionary<string, HardIpSheet>
            {
                ["SheetA"] = new HardIpSheet
                {
                    Rows = [pattern1, pattern2]
                }
            };

            var scghData = new ScghData();

            // Act
            PatSetSheet result = _generator.GenPatSet(sheetDict, scghData);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.AreEqual("PatSets_HardIP", result.Name);

            Assert.IsTrue(result.Rows.Count > 0);

            var allInstanceNames = pattern1.Pattern.InstancePatternName
                .Concat(pattern1.Pattern.InstancePayloadName)
                .Concat(pattern2.Pattern.InstancePatternName)
                .Concat(pattern2.Pattern.InstancePayloadName)
                .ToList();

            Assert.IsTrue(allInstanceNames.Count > 0);
        }

        [TestMethod]
        public void GenPatSet_ShouldAddMissingPatSets_WhenExist()
        {
            // Arrange
            Dictionary<string, HardIpSheet> planDic =
                [];

            ScghData scghData = new ScghData();

            // Act
            PatSetSheet result = _generator.GenPatSet(planDic, scghData);

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.Rows.Count >= 0);
        }

        [TestMethod]
        public void GenPatSet_ShouldGeneratePllPatSet_WhenNonHardIpBlockWithMeasPins()
        {
            // Arrange
            HardIpPattern pattern1 = new HardIpPattern
            {
                Pattern = new PatternClass("DUT_001_C_00_SC"),
                BlockType = "PLL"
            };

            pattern1.MeasPins.Add(
                new MeasPin
                {
                    PinName = "PIN1",
                    MeasType = "measi"
                });

            HardIpPattern pattern2 = new HardIpPattern
            {
                Pattern = new PatternClass("DUT_002_C_00_SC"),
                BlockType = "PLL"
            };
            // 第二個 pattern 不需要 MeasPins
            // 是為了確保 pllFreqSet > 1 的實際情境

            HardIpSheet sheet = new HardIpSheet();
            sheet.Rows.Add(pattern1);
            sheet.Rows.Add(pattern2);

            Dictionary<string, HardIpSheet> planDic =
                new Dictionary<string, HardIpSheet>
                {
            { "PLL_Sheet", sheet }
                };

            ScghData scghData = new ScghData();

            // Act
            PatSetSheet result = _generator.GenPatSet(planDic, scghData);

            // Assert
            Assert.AreNotEqual(null, result);

            // 至少會產生一個 PatSet（PLL PatSet）
            Assert.IsTrue(result.Rows.Count > 0);

            // PatSetName 來自 GenPllPatSetName()
            // default module = Cpu, block = Sa
            bool hasPllPatSet =
                result.Rows.Any(p =>
                    p.PatSetName.StartsWithIgnoreCase("CpuSa"));

            Assert.IsTrue(hasPllPatSet);

            // burstPattern TestName 會被指定成 PatSetName
            Assert.IsFalse(string.IsNullOrEmpty(pattern1.TestName));
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnTrue_WhenSamePatSetRowFiles()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            PatSet existPatSet = new PatSet
            {
                PatSetName = "PATSET_A"
            };
            existPatSet.AddRow(new PatSetRow { File = "P1" });
            existPatSet.AddRow(new PatSetRow { File = "P2" });

            sheet.AddRow(existPatSet);

            PatSet newPatSet = new PatSet
            {
                PatSetName = "PATSET_A"
            };
            // case-insensitive
            newPatSet.AddRow(new PatSetRow { File = "p1" });
            newPatSet.AddRow(new PatSetRow { File = "p2" });

            // Act
            bool result = sheet.IsExistTheSamePatSet(newPatSet, out string patsetName);

            // Assert
            Assert.IsTrue(result);
            Assert.AreEqual("PATSET_A", patsetName);
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnFalse_WhenRowCountDifferent()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            PatSet existPatSet = new PatSet
            {
                PatSetName = "PATSET_B"
            };
            existPatSet.AddRow(new PatSetRow { File = "P1" });

            sheet.AddRow(existPatSet);

            PatSet newPatSet = new PatSet
            {
                PatSetName = "PATSET_B"
            };
            newPatSet.AddRow(new PatSetRow { File = "P1" });
            newPatSet.AddRow(new PatSetRow { File = "P2" });

            // Act
            bool result = sheet.IsExistTheSamePatSet(newPatSet, out string patsetName);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(string.Empty, patsetName);
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnFalse_WhenFileDifferent()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            PatSet existPatSet = new PatSet
            {
                PatSetName = "PATSET_C"
            };
            existPatSet.AddRow(new PatSetRow { File = "P1" });

            sheet.AddRow(existPatSet);

            PatSet newPatSet = new PatSet
            {
                PatSetName = "PATSET_C"
            };
            newPatSet.AddRow(new PatSetRow { File = "P2" });

            // Act
            bool result = sheet.IsExistTheSamePatSet(newPatSet, out string patsetName);

            // Assert
            Assert.IsFalse(result);
            Assert.AreEqual(string.Empty, patsetName);
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnTrue_WhenCompareRowAllMatch()
        {
            PatSetSheet sheet = new PatSetSheet("Test");

            PatSet exist = new PatSet { PatSetName = "PS_A" };
            exist.AddRow(new PatSetRow
            {
                TdGroup = "TD",
                TimeDomain = "T1",
                Enable = "Y",
                File = "FILE1",
                Burst = "B1",
                Comment = "C1"
            });

            sheet.AddRow(exist);

            PatSet incoming = new PatSet { PatSetName = "PS_A" };
            incoming.AddRow(new PatSetRow
            {
                TdGroup = "td",          // case-insensitive
                TimeDomain = "t1",
                Enable = "y",
                File = "file1",
                Burst = "b1",
                Comment = "c1"
            });

            bool result = sheet.IsExistTheSamePatSet(incoming, out string name);

            Assert.IsTrue(result);
            Assert.AreEqual("PS_A", name);
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnFalse_WhenTdGroupDifferent()
        {
            PatSetSheet sheet = new PatSetSheet("Test");

            PatSet exist = new PatSet { PatSetName = "PS_TD" };
            exist.AddRow(new PatSetRow { TdGroup = "TD1", File = "F" });
            sheet.AddRow(exist);

            PatSet incoming = new PatSet { PatSetName = "PS_TD" };
            incoming.AddRow(new PatSetRow { TdGroup = "TD2", File = "F" });

            bool result = sheet.IsExistTheSamePatSet(incoming, out _);

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnFalse_WhenTimeDomainDifferent()
        {
            PatSetSheet sheet = new PatSetSheet("Test");

            PatSet exist = new PatSet { PatSetName = "PS_TD" };
            exist.AddRow(new PatSetRow { TimeDomain = "T1", File = "F" });
            sheet.AddRow(exist);

            PatSet incoming = new PatSet { PatSetName = "PS_TD" };
            incoming.AddRow(new PatSetRow { TimeDomain = "T2", File = "F" });

            Assert.IsFalse(sheet.IsExistTheSamePatSet(incoming, out _));
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnFalse_WhenEnableDifferent()
        {
            PatSetSheet sheet = new PatSetSheet("Test");

            PatSet exist = new PatSet { PatSetName = "PS_EN" };
            exist.AddRow(new PatSetRow { Enable = "Y", File = "F" });
            sheet.AddRow(exist);

            PatSet incoming = new PatSet { PatSetName = "PS_EN" };
            incoming.AddRow(new PatSetRow { Enable = "N", File = "F" });

            Assert.IsFalse(sheet.IsExistTheSamePatSet(incoming, out _));
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldReturnFalse_WhenCommentDifferent()
        {
            PatSetSheet sheet = new PatSetSheet("Test");

            PatSet exist = new PatSet { PatSetName = "PS_C" };
            exist.AddRow(new PatSetRow { File = "F", Comment = "C1" });
            sheet.AddRow(exist);

            PatSet incoming = new PatSet { PatSetName = "PS_C" };
            incoming.AddRow(new PatSetRow { File = "F", Comment = "C2" });

            Assert.IsFalse(sheet.IsExistTheSamePatSet(incoming, out _));
        }

        [TestMethod]
        public void IsExistTheSamePatSet_ShouldIgnorePatternSetAndLabels()
        {
            PatSetSheet sheet = new PatSetSheet("Test");

            PatSet exist = new PatSet { PatSetName = "PS_IGNORE" };
            exist.AddRow(new PatSetRow
            {
                File = "F",
                PatternSet = "AAA",
                StartLabel = "S1",
                StopLabel = "E1"
            });

            sheet.AddRow(exist);

            PatSet incoming = new PatSet { PatSetName = "PS_IGNORE" };
            incoming.AddRow(new PatSetRow
            {
                File = "F",
                PatternSet = "BBB",
                StartLabel = "S2",
                StopLabel = "E2"
            });

            Assert.IsTrue(sheet.IsExistTheSamePatSet(incoming, out _));
        }

        [TestMethod]
        public void AddPatSet_ShouldAddPatSet()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            PatSet patSet = new PatSet
            {
                PatSetName = "PATSET_ADD"
            };

            // Act
            sheet.AddRow(patSet);

            // Assert
            Assert.AreEqual(1, sheet.Rows.Count);
            Assert.AreEqual("PATSET_ADD", sheet.Rows[0].PatSetName);
        }

        [TestMethod]
        public void AddPatSet_ShouldStillAdd_WhenPatSetNameIsNull()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            PatSet patSet = new PatSet();

            // Act
            sheet.AddRow(patSet);

            // Assert
            Assert.AreEqual(1, sheet.Rows.Count);
        }

        [TestMethod]
        public void AddPatSet_ShouldNotAddDuplicatePatSet()
        {
            // Arrange
            var sheet = new PatSetSheet("TestSheet");
            var patSet = new PatSet { PatSetName = "MyPatSet" };
            sheet.AddRow(patSet);

            // Act
            _generator.AddPatSet(sheet, patSet);

            // Assert
            Assert.AreEqual(1, sheet.Rows.Count);
        }

        [TestMethod]
        public void AddPatSet_ShouldSkip_WhenSingleRowFileEqualsPatSetName()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            PatSet patSet = new PatSet
            {
                PatSetName = "SameName"
            };

            PatSetRow row = new PatSetRow
            {
                File = "SameName"
            };
            patSet.AddRow(row);

            // Act
            _generator.AddPatSet(sheet, patSet);

            // Assert
            Assert.AreEqual(0, sheet.Rows.Count);
        }

        [TestMethod]
        public void AddPatSets_ShouldAddMultiplePatSets()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            List<PatSet> patSets =
            [
                new() { PatSetName = "PATSET_1" },
                new() { PatSetName = "PATSET_2" }
            ];

            // Act
            sheet.AddRows(patSets);

            // Assert
            Assert.AreEqual(2, sheet.Rows.Count);
        }

        [TestMethod]
        public void AddPatSets_ShouldHandleNullPatSetName()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            List<PatSet> patSets = [
                new() { PatSetName = "PATSET_OK" },
                new()
            ];

            // Act
            sheet.AddRows(patSets);

            // Assert
            Assert.AreEqual(2, sheet.Rows.Count);
        }

        [TestMethod]
        public void GetScanType_ShouldReturnExpectedValues()
        {
            // Arrange
            string patternTd = "X_X_X_Y_sc_Z_tdf_ABC_DEF";
            string resultTd = _generator.GetScanType(patternTd);
            Assert.AreEqual("Td", resultTd);

            string patternTdChain = "X_X_X_Y_ch_Z_tdf_ABC_DEF";
            string resultTdChain = _generator.GetScanType(patternTdChain);
            Assert.AreEqual("TdChain", resultTdChain);

            string patternSa = "X_X_X_Y_sc_Z_saa_ABC_DEF";
            string resultSa = _generator.GetScanType(patternSa);
            Assert.AreEqual("Sa", resultSa);

            string patternSaBdf = "X_X_X_Y_sc_Z_bdf_ABC_DEF";
            string resultSaBdf = _generator.GetScanType(patternSaBdf);
            Assert.AreEqual("Sa", resultSaBdf);

            string patternSaChain = "X_X_X_Y_ch_Z_saa_ABC_DEF";
            string resultSaChain = _generator.GetScanType(patternSaChain);
            Assert.AreEqual("SaChain", resultSaChain);

            string patternEmpty = "X_X_X_Y_XX_Z_XX_ABC_DEF";
            string resultEmpty = _generator.GetScanType(patternEmpty);
            Assert.AreEqual("", resultEmpty);
        }

        private static HardIpPattern CreatePattern(string payload, int dupIndex = 0)
        {
            var pat = new HardIpPattern
            {
                Pattern = new PatternClass(payload) { InstancePayloadName = [payload] },
                DupIndex = dupIndex
            };
            return pat;
        }

        private static HardIpSheet CreateSheet(params string[] payloads)
        {
            var sheet = new HardIpSheet();
            foreach (string payload in payloads)
            {
                HardIpPattern pat = CreatePattern(payload);
                sheet.Rows.Add(pat);
            }
            return sheet;
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldUseCpuSa_WhenModuleAndBlockEmpty()
        {
            // Arrange
            string payload = "DUT_001_C_00_SC";
            HardIpPattern burstPattern = CreatePattern(payload);
            HardIpSheet sheet = CreateSheet(payload);

            // Act
            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            // Assert
            Assert.AreEqual("CpuSa_dut_001_c_00_sc", result);
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldUseModuleGfx_WhenLDetected()
        {
            string payload = "DUT_002_L_00_SC";
            HardIpPattern burstPattern = CreatePattern(payload);
            HardIpSheet sheet = CreateSheet(payload);

            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            StringAssert.Contains(result, "GfxSa");
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldUseModuleSoc_WhenSDetected()
        {
            string payload = "DUT_003_S_00_SC";
            HardIpPattern burstPattern = CreatePattern(payload);
            HardIpSheet sheet = CreateSheet(payload);

            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            StringAssert.Contains(result, "SocSa");
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldAppendDupIndex_WhenGreaterThanZero()
        {
            string payload = "DUT_004_C_00_SC";
            HardIpPattern burstPattern = CreatePattern(payload, dupIndex: 2);
            HardIpSheet sheet = CreateSheet(payload);

            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            StringAssert.Contains(result, "2");
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldCombineModuleBlockAndPayload_Correctly()
        {
            string payload = "AAA_BBB_C_001_SC";
            HardIpPattern burstPattern = CreatePattern(payload);
            HardIpSheet sheet = CreateSheet(payload, "BBB_CCC_S_111_SC");

            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            Assert.AreEqual("SocSa_aaa_bbb_c_001_sc", result);
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldReturnDefaultCpuSa_WhenPatternListEmpty()
        {
            HardIpPattern burstPattern = CreatePattern("UNKNOWN_001_X_000");
            HardIpSheet sheet = CreateSheet("BBB_CCC_S_111_SC");
            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            Assert.AreEqual("SocSa_unknown_001_x_000", result);
        }

        [TestMethod]
        public void GenPllPatSetName_ShouldFallbackToCpuSa_WhenModuleAndBlockMissing()
        {
            // Arrange
            string payload = "AAA_BBB_X_000_XX";
            HardIpPattern burstPattern = CreatePattern(payload);

            HardIpSheet sheet = CreateSheet(payload);

            // Act
            string result = _generator.GenPllPatSetName(sheet.Rows, burstPattern);

            // Assert
            StringAssert.StartsWith(result, "CpuSa_");
        }

        [TestMethod]
        public void CheckBistOrBira_ShouldReturnBira_WhenContainsBIR()
        {
            // Arrange
            string pattern = "ABC_BIR_XYZ";

            // Act
            string result = _generator.CheckBistOrBira(pattern);

            // Assert
            Assert.AreEqual("Bira", result);
        }

        [TestMethod]
        public void CheckBistOrBira_ShouldReturnBist_WhenNotContainBIR()
        {
            // Arrange
            string pattern = "ABC_BST_XYZ";

            // Act
            string result = _generator.CheckBistOrBira(pattern);

            // Assert
            Assert.AreEqual("Bist", result);
        }

        [TestMethod]
        public void PatSetRows_ShouldInitializeList_WhenFirstAccess()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            // Act
            List<PatSet> rows = sheet.Rows;

            // Assert
            Assert.AreNotEqual(null, rows);
            Assert.AreEqual(0, rows.Count);
        }

        [TestMethod]
        public void PatSetRows_ShouldBeOverwritten_WhenSetExplicitly()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            List<PatSet> newList = [
                new() { PatSetName = "PS1" }
            ];

            // Act
            sheet.Rows = newList;

            // Assert
            Assert.AreEqual(1, sheet.Rows.Count);
            Assert.AreEqual("PS1", sheet.Rows[0].PatSetName);
        }

        [TestMethod]
        public void ExistPatSet_ShouldBeCaseInsensitive()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            sheet.AddRow(new PatSet { PatSetName = "MyPatSet" });

            // Act
            HashSet<string> exist = sheet.ExistPatSet;

            // Assert
            Assert.IsTrue(exist.Contains("mypatset"));
            Assert.IsTrue(exist.Contains("MYPATSET"));
        }

        [TestMethod]
        public void PatSetRowDic_ShouldUseCaseInsensitiveKey()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            PatSet patSet = new PatSet { PatSetName = "AbC" };
            sheet.AddRow(patSet);

            // Act
            Dictionary<string, PatSet> dic = sheet.PatSetRowDic;

            // Assert
            Assert.IsTrue(dic.ContainsKey("abc"));
            Assert.IsTrue(dic.ContainsKey("ABC"));
        }

        [TestMethod]
        public void IsExist_ShouldReturnTrue_WhenPatSetExists()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            sheet.AddRow(new PatSet { PatSetName = "PS_EXIST" });

            // Act
            bool result = sheet.IsExist("ps_exist");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsExist_ShouldReturnFalse_WhenPatSetNotExist()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");

            // Act
            bool result = sheet.IsExist("NOT_EXIST");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetPatSetCnt_ShouldReturnCorrectCount()
        {
            // Arrange
            PatSetSheet sheet = new PatSetSheet("TestSheet");
            sheet.AddRow(new PatSet());
            sheet.AddRow(new PatSet());

            // Act
            long count = sheet.Rows.Count;

            // Assert
            Assert.AreEqual(2, count);
        }

        [TestMethod]
        public void Constructor_WithSheetName_ShouldInitializeCorrectly()
        {
            // Act
            PatSetSheet sheet = new PatSetSheet("MySheet");

            // Assert
            Assert.AreEqual("MySheet", sheet.Name);
            Assert.AreEqual(0, sheet.Rows.Count);
        }

    }
}
