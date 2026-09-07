using System.Collections.Generic;

using Automation.Reader.ScghFile.ProCharPatternSet.Base;
using Automation.Reader.ScghFile.ProCharPatternSet.Business;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScghLib.Base;
using ScghLib.Reader;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class ProdCharPatSetConstructorBaseTests
    {
        private List<ProdCharSheetRow> _rows = null!;

        [TestInitialize]
        public void Setup()
        {
            _rows =
            [
                new()
                {
                    PatternList = ["IN_INIT_01_TEST"],
                    Item = "Item1",
                    Mode = "Mode1",
                    RowNum = 1
                },
                new()
                {
                    PatternList = ["PAYLOAD_01_TEST"],
                    Item = "Item2",
                    Mode = "Mode2",
                    RowNum = 2
                }
            ];
        }

        [TestMethod]
        public void IsInitPattern_ShouldReturnTrue_ForValidInitPattern()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("IN_X_X_INIT_TEST");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsInitPattern_ShouldReturnFalse_ForNonInitPattern()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("PAYLOAD_01_TEST");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetSubName_ShouldReturnCorrectSubstring()
        {
            string result = new HardIpPatSetConstructor(_rows, new List<string>()).GetSubName("A_B_C_D", "0,2");
            Assert.AreEqual("A_C", result);
        }

        [TestMethod]
        public void GetMode_ShouldReturnEmpty_WhenPatternHasNoMode()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            string result = constructor.GetMode("IN_INIT_01_TEST");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetPatSetFromProdChar_ShouldGeneratePatternSet()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            List<ProdCharPatternSetBase> patSets = constructor.GetPatSetFromProdChar(
                new List<ProdCharSheetRow> { _rows[0] },
                new List<ProdCharSheetRow> { _rows[1] }
            );
            Assert.AreEqual(1, patSets.Count);
            Assert.AreEqual(_rows[1].RowNum, patSets[0].RowNum);
        }

        #region IsInitPattern

        [TestMethod]
        public void IsInitPattern_FewerThanFourTokens_ReturnsFalse()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("A_B_C");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsInitPattern_FourToSixTokensNoInMatch_ReturnsFalse()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("A_B_C_D_E");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsInitPattern_GfxSpc0Pattern_ReturnsTrue()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("A_B_L_D_E_F_SPC");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsInitPattern_RscrClearFullPattern_ReturnsTrue()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("XL_FULP_BI_CL00_EFU_JTG_REP_ALLFV_SI");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsInitPattern_RboxRestoreFullPattern_ReturnsTrue()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("XL_FULP_BI_CL00_BHR_JTG_UNS_ALLFV_SI");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsInitPattern_NoBranchMatches_ReturnsFalse()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("AA_BB_CC_DD_EE_FF_GG_HH");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsInitPattern_GfxLMatchesButSpcTokenDoesNotMatch_ReturnsFalse()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("A_B_L_D_E_F_XXXX");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsInitPattern_SpcTokenMatchesButLTokenDoesNot_ReturnsFalse()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("A_B_X_D_E_F_SPC");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsInitPattern_ExactlyFourTokensMatchingIn_ReturnsTrue()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            bool result = constructor.IsInitPattern("A_B_C_INIT");
            Assert.IsTrue(result);
        }

        #endregion

        #region Constructor init/payload split

        [TestMethod]
        public void Constructor_PopulatesInitListAndPayloadListBasedOnIsInitPattern()
        {
            var initRow = new ProdCharSheetRow { RowNum = 1, PayloadList = ["A_B_L_D_E_F_SPC"] };
            var payloadRow = new ProdCharSheetRow { RowNum = 2, PayloadList = ["NOT_AN_INIT_PATTERN"] };
            var rows = new List<ProdCharSheetRow> { initRow, payloadRow };

            var constructor = new HardIpPatSetConstructor(rows, new List<string>());

            List<IProdCharItem> initList = [.. constructor.InitList];
            List<IProdCharItem> payloadList = [.. constructor.PayloadList];

            Assert.AreEqual(1, initList.Count);
            Assert.AreEqual(1, initList[0].RowNum);
            Assert.AreEqual(1, payloadList.Count);
            Assert.AreEqual(2, payloadList[0].RowNum);
        }

        #endregion

        #region GetPatSetFromProdChar init resolution via combined row list

        [TestMethod]
        public void GetPatSetFromProdChar_ResolvesInitByModeLookupAcrossInitAndPayloadRows()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var initSourceRow = new ProdCharSheetRow
            {
                RowNum = 11,
                Mode = "ModeA",
                PayloadList = ["MATCHED_INIT_PATTERN_NAME_X_Y_Z"]
            };
            var dataRow = new ProdCharSheetRow
            {
                RowNum = 20,
                InitList = ["ModeA"],
                PayloadList = ["CARRIER_PAYLOAD_VAL"]
            };

            List<ProdCharPatternSetBase> result = constructor.GetPatSetFromProdChar(
                new List<ProdCharSheetRow> { initSourceRow },
                new List<ProdCharSheetRow> { dataRow }
            );

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("MATCHED_INIT_PATTERN_NAME_X_Y_Z", result[0].InitList[0].PatternName);
            Assert.IsTrue(result[0].InitAliasList.Contains("ModeA"));
        }

        [TestMethod]
        public void GetPatSetFromProdChar_InitTokenIsNa_TreatedAsEmptyPlaceholderWithoutLookup()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var initSourceRow = new ProdCharSheetRow { RowNum = 30, PayloadList = ["DUMMY_INIT_SOURCE"] };
            var dataRow = new ProdCharSheetRow { RowNum = 31, InitList = ["NA"], PayloadList = ["CARRIER_VAL"] };

            List<ProdCharPatternSetBase> result = constructor.GetPatSetFromProdChar(
                new List<ProdCharSheetRow> { initSourceRow },
                new List<ProdCharSheetRow> { dataRow }
            );

            // The resolved "" init pattern name is later stripped by DeleteSamePattern (empty names are dropped).
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(0, result[0].InitList.Count);
        }

        [TestMethod]
        public void GetPatSetFromProdChar_RawEmptyInitToken_TreatedAsEmptyPlaceholderWithoutLookup()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var initSourceRow = new RawProdCharItem { RowNum = 40, Mode = "UNRELATED_MODE" };
            var dataRow = new RawProdCharItem
            {
                RowNum = 41,
                RawInitList = [""],
                PayloadValue = "CARRIER_MARKER_VALUE"
            };

            List<ProdCharPatternSetBase> result = constructor.GetPatSetFromProdChar(
                [initSourceRow],
                [dataRow]
            );

            // The resolved "" init pattern name is later stripped by DeleteSamePattern (empty names are dropped).
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(0, result[0].InitList.Count);
        }

        #endregion

        #region GetMode

        [TestMethod]
        public void GetMode_TenthTokenMatchesPerformanceModePattern_ReturnsToken()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            string result = constructor.GetMode("A_B_C_D_E_F_G_H_I_Mabc123x");
            Assert.AreEqual("Mabc123x", result);
        }

        [TestMethod]
        public void GetMode_TenthTokenContains999_ReturnsEmpty()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            string result = constructor.GetMode("A_B_C_D_E_F_G_H_I_Mabc999x");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetMode_TenthTokenContains010_ReturnsEmpty()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            string result = constructor.GetMode("A_B_C_D_E_F_G_H_I_M010abc");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetMode_TenthTokenDoesNotMatchPattern_ReturnsEmpty()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            string result = constructor.GetMode("A_B_C_D_E_F_G_H_I_abc123");
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void GetMode_ExactlyNineTokens_ReturnsEmpty()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            string result = constructor.GetMode("A_B_C_D_E_F_G_H_Mabc123x");
            Assert.AreEqual("", result);
        }

        #endregion

        // Minimal IProdCharItem implementation that returns its raw lists verbatim, unlike
        // ProdCharSheetRow.GetPayloadList()/GetInitList() which pre-filter blank entries.
        private sealed class RawProdCharItem : IProdCharItem
        {
            public int RowNum { get; set; }
            public string PayloadValue { get; set; } = "";
            public string Block { get; set; } = "";
            public string Mode { get; set; } = "";
            public string Item { get; set; } = "";
            public string Usage { get; set; } = "";
            public string Inits { get; set; } = "";
            public string PayLoads { get; set; } = "";
            public string Chiplet { get; set; } = "";
            public List<string> RawInitList { get; set; } = [];
            public List<string> RawPayloadList { get; set; } = [];
            public string SourceSheet { get; set; } = "";

            public List<string> GetInitList() => RawInitList;
            public List<string> GetPayloadList() => RawPayloadList;
            public string GetSourceSheet() => SourceSheet;
        }

        #region FilterProChar

        [TestMethod]
        public void FilterProChar_UsageZero_IsExcluded()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var inputList = new List<ProdCharSheetRow>
            {
                new() { RowNum = 1, Usage = "0" },
                new() { RowNum = 2, Usage = "1" }
            };

            List<IProdCharItem> result = constructor.FilterProChar(inputList);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(2, result[0].RowNum);
        }

        #endregion

        #region GetPayloadList

        [TestMethod]
        public void GetPayloadList_RawEmptyStringPayload_AddsEmptyStringWithoutLookup()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var row = new RawProdCharItem { RawPayloadList = [""] };

            List<string> result = constructor.GetPayloadList(row, []);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("", result[0]);
        }

        [TestMethod]
        public void GetPayloadList_PayloadIsNa_AddsEmptyStringWithoutLookup()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var row = new RawProdCharItem { RawPayloadList = ["NA"] };

            List<string> result = constructor.GetPayloadList(row, []);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("", result[0]);
        }

        [TestMethod]
        public void GetPayloadList_PayloadIsNSlashA_AddsEmptyStringWithoutLookup()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var row = new RawProdCharItem { RawPayloadList = ["N/A"] };

            List<string> result = constructor.GetPayloadList(row, []);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("", result[0]);
        }

        [TestMethod]
        public void GetPayloadList_NoMatchFound_FallsBackToCurrentPayloadAndRecordsPayLoads()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var row = new RawProdCharItem { RawPayloadList = ["REAL_VALUE"] };

            List<string> result = constructor.GetPayloadList(row, []);

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("REAL_VALUE", result[0]);
            Assert.AreEqual("REAL_VALUE,", row.PayLoads);
        }

        #endregion

        #region RemoveSamePatternsHashSet

        [TestMethod]
        public void RemoveSamePatternsHashSet_SameReferenceTwice_CollapsesToOne()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var patset = new ProdCharPatternSetBase { RowNum = 1 };
            var patsetList = new List<ProdCharPatternSetBase> { patset, patset };

            List<ProdCharPatternSetBase> result = constructor.RemoveSamePatternsHashSet(patsetList);

            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void RemoveSamePatternsHashSet_DistinctInstancesWithSameValues_AreNotCollapsed()
        {
            // No Equals/GetHashCode override on ProdCharPatternSetBase, so dedup is by reference only.
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var patsetList = new List<ProdCharPatternSetBase>
            {
                new() { RowNum = 1 },
                new() { RowNum = 1 }
            };

            List<ProdCharPatternSetBase> result = constructor.RemoveSamePatternsHashSet(patsetList);

            Assert.AreEqual(2, result.Count);
        }

        #endregion

        #region DeleteSamePattern

        [TestMethod]
        public void DeleteSamePattern_ConsecutiveSameInitPatternName_CollapsesToOne()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var patset = new ProdCharPatternSetBase
            {
                InitList = new Dictionary<int, PatternWithMode>
                {
                    { 0, new PatternWithMode { PatternName = "INIT_A" } },
                    { 1, new PatternWithMode { PatternName = "INIT_A" } }
                },
                PayloadList = []
            };

            constructor.DeleteSamePattern([patset]);

            Assert.AreEqual(1, patset.InitList.Count);
        }

        [TestMethod]
        public void DeleteSamePattern_EmptyInitPatternName_IsSkipped()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var patset = new ProdCharPatternSetBase
            {
                InitList = new Dictionary<int, PatternWithMode>
                {
                    { 0, new PatternWithMode { PatternName = "" } },
                    { 1, new PatternWithMode { PatternName = "INIT_A" } }
                },
                PayloadList = []
            };

            constructor.DeleteSamePattern([patset]);

            Assert.AreEqual(1, patset.InitList.Count);
            Assert.AreEqual("INIT_A", patset.InitList[1].PatternName);
        }

        [TestMethod]
        public void DeleteSamePattern_SamePayloadReferenceTwice_CollapsesToOne()
        {
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var sharedPayload = new PatternWithMode { PatternName = "A_B_C_PAY_A" };
            var patset = new ProdCharPatternSetBase
            {
                InitList = [],
                PayloadList = [sharedPayload, sharedPayload]
            };

            constructor.DeleteSamePattern([patset]);

            Assert.AreEqual(1, patset.PayloadList.Count);
        }

        [TestMethod]
        public void DeleteSamePattern_DistinctPayloadInstancesWithSameName_AreNotCollapsed()
        {
            // PayloadList dedup uses reference equality (no Equals override), unlike InitList which compares PatternName strings.
            var constructor = new HardIpPatSetConstructor(_rows, new List<string>());
            var patset = new ProdCharPatternSetBase
            {
                InitList = [],
                PayloadList =
                [
                    new() { PatternName = "A_B_C_PAY_A" },
                    new() { PatternName = "A_B_C_PAY_A" }
                ]
            };

            constructor.DeleteSamePattern([patset]);

            Assert.AreEqual(2, patset.PayloadList.Count);
        }

        #endregion
    }
}
