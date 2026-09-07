using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Utility.TpUpdate.HardIPEnableWordsUpdate;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Newtonsoft.Json;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class TtrToolWriterTests : FunctionTestBase
    {
        private TtrToolWriter _writer = null!;
        private EnableWordTable _table = null!;
        private Dictionary<string, string> _subflowStatus = null!;
        private List<string> _nonUsedItems = null!;

        [TestInitialize]
        public void Setup()
        {
            _writer = new TtrToolWriter();
            _table = new EnableWordTable();
            _subflowStatus = [];
            _nonUsedItems = [];
        }

        #region Integration style tests (file compare)

        [TestMethod]
        public void UpdateOverlayFlowSheetEnableWords()
        {
            string subName = "UpdateOverlayFlowSheetEnableWords";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            _ = Directory.CreateDirectory(outputPath);

            SubFlowSheet flowSheet = new SubFlowSheet("Test")
            {
                Rows = new FlowRows
                {
                    new FlowRow
                    {
                        SheetName = "Flow_WALKINGZ_3",
                        Opcode = OpCode.Test,
                        Parameter = "relayon",
                    },
                    new FlowRow
                    {
                        SheetName = "Flow_WALKINGZ_3",
                        Opcode = OpCode.Test,
                        Parameter = "WALKINGZ_DC_CONTINUITY_DIFF_NEG_P_25C_FT1",
                        FailAction = "F_WALKINGZ_3",
                        Job = "CP2",
                        TName = "CONTINUITYNEG",
                        Env = "Env"
                    },
                    new FlowRow
                    {
                        Opcode = OpCode.If,
                        Parameter = "flag",
                    },
                    new FlowRow
                    {
                        Opcode = OpCode.For,
                        Parameter = "flag",
                    },
                    new FlowRow
                    {
                        Opcode = OpCode.Next,
                        Parameter = "flag",
                    },
                }
            };

            EnableWordTable ampTable = new EnableWordTable();
            List<string> enableList = [];
            Dictionary<string, string> tableUpdateStatus = [];
            List<string> nonUsedItems = [];
            List<string> tpJobs = [];

            SubFlowSheet result = _writer.UpdateOverlayFlowSheetEnableWords(flowSheet, ampTable, enableList, tableUpdateStatus, ref nonUsedItems, tpJobs);

            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(Path.Combine(outputPath, "result.json"), json);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void FailControl_ShouldBeGenerated_WhenPreviousTestExists()
        {

            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_FAILCTRL_LV",
                    FailAction = "FC_FLAG"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_BLK_PAT_A_LV"
                });

            _table.Rows = new List<EnableWordTableRow>
            {
                CreateEnableWordRow(
                    "PAT_A",
                    "CAT",
                    "BLK",
                    "FAILCTRL",
                    new Dictionary<string, List<string>>
                    {
                        { "LV", new List<string> { "EN_LV" } }
                    })
            };


            (SubFlowSheet _, List<FailControlData> failControls) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(1, failControls.Count);
            Assert.IsTrue(failControls[0].FailFlagName.Contains("FC_FLAG"));
        }

        [TestMethod]
        public void FailControl_ShouldNotGenerate_WhenNoPreviousTest()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "PAT_A_LV"
                });

            _table.Rows = new List<EnableWordTableRow>
            {
                CreateEnableWordRow(
                    "PAT_A",
                    "CAT",
                    "BLK",
                    "FAILCTRL",
                    new Dictionary<string, List<string>>
                    {
                        { "LV", new List<string> { "EN_LV" } }
                    })
            };

            (SubFlowSheet _, List<FailControlData> failControls) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(0, failControls.Count);
        }

        [TestMethod]
        public void Enable_ShouldContainBVValidation_WhenLvBinCutExists()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "ABC_BLK_PAT_A_LV",
                    Enable = "DUMMY"
                });


            _table.Rows = new List<EnableWordTableRow>
            {
                CreateEnableWordRow(
                    "PAT_A",
                    "CAT",
                    "BLK",
                    "",
                    new Dictionary<string, List<string>>
                    {
                        { "LV", new List<string> { "EN_LV" } },
                        { "BV", new List<string> { "BV_EN" } }
                    })
            };

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.IsTrue(result.Rows[0].Enable.Contains("BV_Validation"));
        }

        [TestMethod]
        public void MergeEnable_ShouldCollapseToAll()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_BLK_PAT_A_LV",
                    Enable = "OLD"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
                    { "LV", new List<string> { "EN_cp1", "EN_cp2" } }
                });

            _table.Rows = new List<EnableWordTableRow> { row };
            _table.EnableWords = new List<string> { "EN_cp1", "EN_cp2" };

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.IsTrue(result.Rows[0].Enable.Contains("EN_All"));
        }

        [TestMethod]
        public void Relay_ShouldBeRemoved_WhenOnlyOnOffPair()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow { Opcode = "Test", Parameter = "AtgRelay_relayon_K1" },
                new FlowRow { Opcode = "Test", Parameter = "AtgRelay_relayoff_K1" }
            );

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(0, result.Rows.Count);
        }


        [TestMethod]
        public void Relay_ShouldPropagateEnable_WhenThreeRows()
        {

            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "relayon_K1"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_BLK_PAT_A_LV",
                    Enable = "EN"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "relayoff_K1"
                });

            _table.Rows = new List<EnableWordTableRow>
            {
                CreateEnableWordRow(
                    "PAT_A",
                    "CAT",
                    "BLK",
                    "",
                    new Dictionary<string, List<string>>
                    {
                        { "LV", new List<string> { "EN_LV" } }
                    })
            };

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.IsTrue(result.Rows.Any(r => r.Enable.Contains("EN_LV")));
        }

        [TestMethod]
        public void Overlay_ShouldApplyEnable_FromOverlayDict()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_MA123A",
                    Enable = "BASE"
                });

            _table.OverlayInfoDic.Add(
                "MA123A",
                new List<string> { "OV_EN" });

            SubFlowSheet result = _writer.UpdateOverlayFlowSheetEnableWords(
                sheet, _table, new List<string>(), _subflowStatus, ref _nonUsedItems);

            Assert.IsTrue(result.Rows[0].Enable.Contains("OV_EN"));
        }

        [TestMethod]
        public void InstanceOverlay_ShouldBeSet()
        {
            InstanceSheet sheet = new InstanceSheet("Test");
            sheet.AddRow(new InstanceRow
            {
                TestName =
                    "XXX_BBB_AAA_BBB_CCC_DDD_EEE_FFF_GGG_HHH_III_MODE_LV"
            });

            _table.Rows = new List<EnableWordTableRow>
            {
                CreateEnableWordRow(
                    "AAA_BBB_CCC_DDD_EEE_FFF_GGG_HHH_III_MODE",
                    "CAT",
                    "BBB",
                    "",
                    new Dictionary<string, List<string>>
                    {
                        { "LV_PBCut", new List<string> { "X" } }
                    })
            };

            var map = new Dictionary<string, string>
            {
                { "MODE", "Overlay1" }
            };

            InstanceSheet result = _writer.UpdateInsSheetOverlays(sheet, _table, map);

            Assert.AreEqual("Overlay1", result.Rows[0].Overlay);
        }

        [TestMethod]
        public void WriteFailControlTable()
        {
            string subName = "WriteFailControlTable";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            _ = Directory.CreateDirectory(outputPath);

            List<FailControlData> failControls =
            [
                new()
                {
                    InstanceName = "InstanceA",
                    FailFlagName = "FailFlagA",
                    EnableAllSite = "TRUE"
                }
            ];

            _writer.WriteFailControlTable(outputPath, failControls);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void WriteFailControlTable_ShouldReturnEmpty_WhenNoData()
        {
            string result =
                _writer.WriteFailControlTable(OutputPath, []);

            Assert.AreEqual("", result);
        }

        #endregion

        #region Logic / branch coverage tests
        [TestMethod]
        public void EnableNotChanged_WhenTableRowNotMatched()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "PAT_X_LV",
                    Enable = "A&&B"
                });

            _table.Rows = [];

            (SubFlowSheet result, List<FailControlData> failControls) = _writer.UpdateFlowSheetEnableWords(sheet, _table, ["A", "B"], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual("SubFlow no change", _subflowStatus["SF"]);
            Assert.AreEqual("A&&B", result.Rows[0].Enable);
            Assert.AreEqual(0, failControls.Count);
        }

        [TestMethod]
        public void EnableCleared_WhenTableRowMatched_ButOriginalEnableRemoved()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "TEST_BLK_PAT_A_LV",
                    Enable = "A&&B"
                });

            _table.Rows =
    [
        CreateEnableWordRow(
            "PAT_A",
            "CAT",
            "BLK",
            "",
            new Dictionary<string, List<string>>
            {
                { "LV", new List<string> { "EN_LV" } }
            })
    ];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, ["A", "B"], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual("SubFlow Edited", _subflowStatus["SF"]);
            Assert.AreEqual("EN_LV", result.Rows[0].Enable);
        }

        [TestMethod]
        public void IfOpcode_AlwaysClearsEnable()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "If",
                    Parameter = "COND",
                    Enable = "A"
                });

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, ["A"], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(string.Empty, result.Rows[0].Enable);
        }

        [TestMethod]
        public void ForNext_PropagatesEnableFromTest()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow { Opcode = "For" },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "TEST_BLK_PAT_A_LV",
                    Enable = "A"
                },
                new FlowRow { Opcode = "Next" });

            _table.Rows =
    [
        CreateEnableWordRow(
            "PAT_A",
            "CAT",
            "BLK",
            "",
            new Dictionary<string, List<string>>
            {
                { "LV", new List<string> { "EN_LV" } }
            })
    ];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, ["A"], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual("EN_LV", result.Rows[0].Enable);
        }

        [TestMethod]
        public void NonUsedItems_ReturnsUnusedRows()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "TEST_BLK_PAT_A_LV"
                });

            EnableWordTableRow usedRow = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            EnableWordTableRow unusedRow = CreateEnableWordRow(
                "PAT_B",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            _table.Rows = [usedRow, unusedRow];

            _writer.UpdateFlowSheetEnableWords(sheet, _table, [], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(1, _nonUsedItems.Count);
            Assert.IsTrue(_nonUsedItems[0].Contains("BLK_PAT_B"));
        }

        [TestMethod]
        public void Match_Other_Fallback_When_No_SubBlock_Match()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_PAT_A_LV",
                    Enable = "A"
                });

            EnableWordTableRow otherRow = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "other",
                "",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN_LV" } }
                });

            _table.Rows = [otherRow];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, ["A"], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual("EN_LV", result.Rows[0].Enable);
        }
        #endregion

        #region Helpers
        private static SubFlowSheet CreateSheet(params FlowRow[] flowRowArray)
        {
            SubFlowSheet sheet = new SubFlowSheet("SF");
            foreach (FlowRow row in flowRowArray)
            {
                sheet.AddRow(row);
            }

            return sheet;
        }

        private static EnableWordTableRow CreateEnableWordRow(string pattern, string category, string subBlock, string failControl, Dictionary<string, List<string>> enableWords)
        {
            Dictionary<string, int> indexMap = new Dictionary<string, int>
            {
                { "Pattern", 1 },
                { "Category", 2 },
                { "SubBlock", 3 },
                { "HIP_FailRun", 4 },
                { "LV", 5 }
            };

            Dictionary<string, TtrColumnMetadata> headerMetadata = new()
            {
                { "Pattern", new("Pattern", 1) },
                { "Category", new("Category", 2) },
                { "SubBlock", new("SubBlock", 3) },
                { "HIP_FailRun", new("HIP_FailRun", 4) },
                { "LV", new("LV", 5) },
            };

            Dictionary<int, string> dataset = new Dictionary<int, string>
            {
                { 1, pattern },
                { 2, category },
                { 3, subBlock },
                { 4, failControl },
                { 5, string.Join(",", enableWords.SelectMany(p => p.Value)) }
            };

            EnableWordTableRow row = new EnableWordTableRow(
                "Pattern",
                "Category",
                "SubBlock",
                "HIP_FailRun",
                headerMetadata,
                dataset)
            {
                EnableWords = enableWords
            };

            return row;
        }
        #endregion

        [TestMethod]
        public void GenerateFailControl_WhenPreviousTestExists_WithFailAction()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AAA_CAT_BLK_FC_LV",
                    FailAction = "FLAG1"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AAA_CAT_BLK_PAT_A_LV"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "FC",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            _table.Rows = [row];

            (_, List<FailControlData> failControls) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, [], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(1, failControls.Count);
            Assert.AreEqual("FLAG1", failControls[0].FailFlagName);
        }

        [TestMethod]
        public void AppendBVValidation_WhenLVContainsBV()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "TEST_BLK_PAT_A_LV",
                    Enable = "A"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
            { "BV", new List<string> { "EN_BV" } },
            { "LV", new List<string> { "EN_LV" } }
                });

            _table.Rows = [row];

            (SubFlowSheet result, _) = _writer.UpdateFlowSheetEnableWords(sheet, _table, [], _subflowStatus, ref _nonUsedItems);

            Assert.IsTrue(result.Rows[0].Enable.Contains("BV_Validation"));
        }

        [TestMethod]
        public void AssignNegativeJobs_WhenNoMatch()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "TEST_BLK_PAT_A_LV",
                    Enable = "A"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            _table.Rows = [row];

            List<string> jobs = ["CP1", "FT1"];

            (SubFlowSheet result, _) = _writer.UpdateFlowSheetEnableWords(sheet, _table, [], _subflowStatus, ref _nonUsedItems, jobs);

            Assert.AreEqual("!CP1,!FT1", result.Rows[0].Job);
        }

        [TestMethod]
        public void CharacterizeOpcode_ShouldBehaveLikeTest()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "characterize",
                    Parameter = "TEST_BLK_PAT_A_LV",
                    Enable = "A"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN_LV" } }
                });

            _table.Rows = [row];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, [], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual("(A||EN_LV)", result.Rows[0].Enable);
        }

        [TestMethod]
        public void Cache_ShouldReuseTableItem()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow { Opcode = "Test", Parameter = "TEST_BLK_PAT_A_LV" },
                new FlowRow { Opcode = "Test", Parameter = "TEST_BLK_PAT_A_LV" });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            _table.Rows = [row];

            _writer.UpdateFlowSheetEnableWords(sheet, _table, [], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual(0, _nonUsedItems.Count);
        }

        [TestMethod]
        public void ClearCurrentEnableWords_ShouldRemoveMatchedEnable()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "TEST_BLK_PAT_X_LV",
                    Enable = "A&&B||C"
                });

            _table.Rows =
            [
                CreateEnableWordRow(
                    "PAT_X",
                    "CAT",
                    "BLK",
                    "",
                    new Dictionary<string, List<string>>
                    {
                        { "LV", new List<string> { "EN" } }
                    })
            ];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(sheet, _table, ["A", "B"], _subflowStatus, ref _nonUsedItems);

            Assert.AreEqual("(C||EN)", result.Rows[0].Enable);
        }

        [TestMethod]
        public void GenerateFailControl_PerTD_ShouldEnableAllSite()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AAA_CAT_BLK_FC_LV",
                    FailAction = "FLAG1"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AAA_CAT_BLK_PAT_A_LV"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "FC,PerTD",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            _table.Rows = [row];

            (_, List<FailControlData> failControls) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet,
                    _table,
                    [],
                    _subflowStatus,
                    ref _nonUsedItems);

            Assert.AreEqual(1, failControls.Count);
            Assert.AreEqual("TRUE", failControls[0].EnableAllSite);
        }

        [TestMethod]
        public void GenerateFailControl_ShouldUseCommentFlag_WhenFailActionEmpty()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AAA_CAT_BLK_FC_LV",
                    FailAction = "",
                    Comment1 = "Dummy\tFLAG_FROM_COMMENT"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AAA_CAT_BLK_PAT_A_LV"
                });

            EnableWordTableRow row = CreateEnableWordRow(
                "PAT_A",
                "CAT",
                "BLK",
                "FC",
                new Dictionary<string, List<string>>
                {
            { "LV", new List<string> { "EN" } }
                });

            _table.Rows = [row];

            (_, List<FailControlData> failControls) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet,
                    _table,
                    [],
                    _subflowStatus,
                    ref _nonUsedItems);

            Assert.AreEqual(1, failControls.Count);
            Assert.AreEqual("FLAG_FROM_COMMENT", failControls[0].FailFlagName);
        }

        [TestMethod]
        public void Enable_ShouldContainBVValidation_WhenHvBinCutExists()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "ABC_BLK_PAT_A_HV",
                    Enable = "DUMMY"
                });

            _table.Rows =
            [
                CreateEnableWordRow(
            "PAT_A",
            "CAT",
            "BLK",
            "",
            new Dictionary<string, List<string>>
            {
                { "HV", new List<string> { "EN_HV" } },
                { "HBV", new List<string> { "HBV_EN" } }
            })
            ];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet,
                    _table,
                    [],
                    _subflowStatus,
                    ref _nonUsedItems);

            Assert.IsTrue(result.Rows[0].Enable.Contains("BV_Validation"));
        }

        [TestMethod]
        public void RelayEnable_ShouldBeCleared_WhenRelaySettingMismatch()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AtgRelay_relayon_K1"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_BLK_PAT_A_LV",
                    Enable = "EN"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AtgRelay_relayoff_K2"
                });

            _table.Rows =
            [
                CreateEnableWordRow(
            "PAT_A",
            "CAT",
            "BLK",
            "",
            new Dictionary<string, List<string>>
            {
                { "LV", new List<string> { "EN_LV" } }
            })
            ];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet,
                    _table,
                    [],
                    _subflowStatus,
                    ref _nonUsedItems);

            Assert.AreEqual(string.Empty, result.Rows[0].Enable);
            Assert.AreEqual(string.Empty, result.Rows[^1].Enable);
        }

        [TestMethod]
        public void RelayEnable_ShouldMergeMultipleEnableWords()
        {
            SubFlowSheet sheet = CreateSheet(
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AtgRelay_relayon_K1"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_BLK_PAT_A_LV",
                    Enable = "AAA"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "XXX_BLK_PAT_B_LV",
                    Enable = "BBB"
                },
                new FlowRow
                {
                    Opcode = "Test",
                    Parameter = "AtgRelay_relayoff_K1"
                });

            _table.Rows =
            [CreateEnableWordRow("PAT_A", "CAT", "BLK", "",
                new Dictionary<string, List<string>>
                {
                    { "LV", new List<string> { "AAA" } }
                }),
                CreateEnableWordRow("PAT_B", "CAT", "BLK", "",
                    new Dictionary<string, List<string>>
                    {
                        { "LV", new List<string> { "BBB" } }
                    })
                    ];

            (SubFlowSheet result, _) =
                _writer.UpdateFlowSheetEnableWords(
                    sheet,
                    _table,
                    [],
                    _subflowStatus,
                    ref _nonUsedItems);

            Assert.AreEqual(result.Rows[0].Enable, result.Rows[^1].Enable);
            Assert.AreNotEqual(string.Empty, result.Rows[0].Enable);
        }
    }
}
