using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Enums;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpFlowRowGeneratorTests : FunctionTestBase
    {
        private HardIpFlowRowGenerator _generator = null!;

        [TestInitialize]
        public void Setup()
        {
            var paraData = new HardIpParaData(EnumBlock.HardIp);
            var inputData = new HardIpInputData(paraData);

            _generator = new HardIpFlowRowGenerator(inputData, "TestSheet")
            {
                Pat = new HardIpPattern
                {
                    SheetName = "TestSheet",
                    Pattern = new PatternClass("PAT_A"),
                    MiscInfo = string.Empty,
                    MiscInfoDict = [],
                    UseLimitsN = [],
                    UseLimitsL = [],
                    UseLimitsH = []
                },

                LabelVoltage = HardIpConstData.LabelNv
            };
            LocalSpecs.Options.Device = EnumDevice.AP;
        }

        #region GenTestRows

        [TestMethod]
        public void GenTestRows_MappingTtrBranch_ShouldUseMappedFailAction()
        {
            _generator.Pat.MiscInfo = $"{HardIpConstData.MappingTtrBranch}:KEY1";
            _generator.InitOriginalItems["KEY1"] = "MAPPED_FLAG";

            List<FlowRow> rows = _generator.GenTestRows();

            FlowRow testRow = rows.First(r => r.Opcode == OpCode.Test);
            Assert.AreEqual("MAPPED_FLAG", testRow.FailAction);
        }

        [TestMethod]
        public void GenTestRows_OriginalTtrBranch_ShouldAppendClearRowsAndOverwriteMap()
        {
            _generator.Pat.MiscInfo = $"{HardIpConstData.OriginalTtrBranch}:ORI_KEY";
            _generator.Pat.Failflag = "NV@FLAG_NV;";

            List<FlowRow> rows = _generator.GenTestRows();

            Assert.IsTrue(rows.Any(r => r.Opcode == OpCode.If));
            Assert.IsTrue(rows.Any(r => r.Opcode == OpCode.EndIf));
            Assert.AreEqual("FLAG_NV", _generator.InitOriginalItems["ORI_KEY"]);
        }

        [TestMethod]
        public void GenTestRows_NoBinOutForAll_ShouldClearFailAction()
        {
            _generator.HardIpInputData.ConfigData.NoBinOutForAll = true;
            _generator.Pat.Failflag = "NV@FAIL_NV;";

            List<FlowRow> rows = _generator.GenTestRows();

            FlowRow testRow = rows.First(r => r.Opcode == OpCode.Test);
            Assert.AreEqual(string.Empty, testRow.FailAction);
            Assert.IsTrue(testRow.Comment1.Contains("FAIL_NV"));
        }
        #endregion

        #region GetIfFlowRows
        [TestMethod]
        public void GetIfFlowRows_MultiFlags_WithDeviceName_ShouldWrapWithIfEndIf()
        {
            FlowRow testRow = new FlowRow
            {
                Opcode = OpCode.Test,
                Job = "JOB1",
                DeviceName = "F_FLAG"
            };

            var limits = new List<FlowRow>
            {
                new() { Opcode = OpCode.Test, Parameter = "LIMIT1" }
            };

            string condition = "F_A&&F_B\nF_C";

            List<FlowRow> rows = _generator.GetIfFlowRows(condition, testRow, limits);

            Assert.AreEqual(OpCode.If, rows.First().Opcode);
            Assert.AreEqual(OpCode.EndIf, rows.Last().Opcode);
            Assert.IsTrue(rows.Any(r => r.Parameter == "LIMIT1"));
        }
        #endregion

        #region GenUseLimitRows
        [TestMethod]
        public void GenUseLimitRows_EmptyUseLimits_ShouldReturnEmpty()
        {
            _generator.Pat.UseLimitsN.Clear();

            List<FlowRow> rows = InvokeGenUseLimitRows("PARA", "FAIL");

            Assert.AreEqual(0, rows.Count);
        }

        [TestMethod]
        public void GenUseLimitRows_NoBinOutForAll_ShouldPutFailIntoComment1()
        {
            _generator.HardIpInputData.ConfigData.NoBinOutForAll = true;
            _generator.Pat.UseLimitsN.Add(new MeasPin
            {
                PinName = "PIN1",
                MeasType = MeasType.MeasV,
                LowLimit = "0",
                HighLimit = "1"
            });

            List<FlowRow> rows = InvokeGenUseLimitRows("PARA", "FAILFLAG");

            Assert.IsTrue(rows.All(r => string.IsNullOrEmpty(r.FailAction)));
            Assert.IsTrue(rows.Any(r => r.Comment1.Contains("FAILFLAG")));
        }
        #endregion

        #region GenShmooRows
        [TestMethod]
        public void GenShmooRows_NoShmoo_ShouldReturnEmpty()
        {
            _generator.Pat.Shmoo = new HardipCharSetup
            {
                CharSteps = [] // 明確空
            };

            List<FlowRow> rows = _generator.GenShmooRows("NV");

            Assert.AreEqual(0, rows.Count);
        }

        [TestMethod]
        public void GenShmooRows_WithShmoo_ShouldGenerateCharacterizeRow()
        {
            _generator.Pat.Shmoo = new HardipCharSetup
            {
                SetupName = "SHMOO1",
                TestNameInFlow = "HARDIP_SHMOO_V",
                CharSteps = [
                    new("SHMOO1", "STEP1")
                ]
            };

            _generator.LabelVoltage = "NV";

            List<FlowRow> rows = _generator.GenShmooRows("NV");

            Assert.AreEqual(1, rows.Count);
            Assert.AreEqual(OpCode.Characterize, rows[0].Opcode);
            Assert.IsTrue(rows[0].Parameter.Contains("SHMOO1"));
        }
        #endregion

        #region GenPreRetestRows
        [TestMethod]
        public void GenPreRetestRows_NoReTest_ShouldReturnEmpty()
        {
            _generator.Pat.MiscInfoDict.Clear();

            List<FlowRow> rows = _generator.GenPreRetestRows();

            Assert.AreEqual(0, rows.Count);
        }

        [TestMethod]
        public void GenPreRetestRows_WithReTest_ShouldGenerateRows()
        {
            _generator.Pat.MiscInfoDict[HardIpConstData.ReTest] = "Y";

            List<FlowRow> rows = _generator.GenPreRetestRows();

            Assert.IsTrue(rows.Any(r => r.FailAction == HardIpConstData.ReTestFlag));
            Assert.IsTrue(rows.Any(r => r.Opcode == OpCode.Test));
        }
        #endregion

        #region GenRetestIfRow
        [TestMethod]
        public void GenRetestIfRow_WithFlag()
        {
            _generator.Pat.MiscInfoDict = new Dictionary<string, string>
            {
                { HardIpConstData.ReTest, "1" }
            };

            FlowRow? row = _generator.GenRetestIfRow();

            Assert.AreNotEqual(null, row);
            Assert.AreEqual("if", row!.Opcode);
        }

        [TestMethod]
        public void GenRetestIfRow_NoFlag_ShouldReturnNull()
        {
            _generator.Pat.MiscInfoDict = [];

            FlowRow? row = _generator.GenRetestIfRow();

            Assert.AreEqual(null, row);
        }
        #endregion

        #region GenPrintStartRow
        [TestMethod]
        public void GenPrintStartRow_Default()
        {
            FlowRow row = _generator.GenPrintStartRow();

            Assert.IsTrue(row.Parameter.Contains("Start"));
        }

        [TestMethod]
        public void GenPrintStopRow_Custom()
        {
            FlowRow row = _generator.GenPrintStopRow("MY_FLOW");

            Assert.IsTrue(row.Parameter.Contains("MY_FLOW"));
        }
        #endregion

        private List<FlowRow> InvokeGenUseLimitRows(string parameter, string failAction)
        {
            return _generator.GenUseLimitRows(parameter, failAction);
        }

        #region GenBinTableRow
        [TestMethod]
        public void GenBinTableRow_Default_BuildsRowFromParts()
        {
            FlowRow row = _generator.GenBinTableRow();

            Assert.AreEqual(OpCode.BinTable, row.Opcode);
        }
        #endregion

        #region CreateIfConditionByTestPanDefine
        [TestMethod]
        public void CreateIfConditionByTestPanDefine_SiteFlagMatchesLabelVoltage_ReturnsCondition()
        {
            _generator.Pat.SiteFlag = "NV@site-var=1;HV@site-var=2;";
            _generator.LabelVoltage = "NV";

            string result = _generator.CreateIfConditionByTestPanDefine();

            Assert.AreEqual("site-var=1", result);
        }

        [TestMethod]
        public void CreateIfConditionByTestPanDefine_NoSiteFlag_ReturnsEmpty()
        {
            _generator.Pat.SiteFlag = "";
            _generator.LabelVoltage = "NV";

            string result = _generator.CreateIfConditionByTestPanDefine();

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void CreateIfConditionByTestPanDefine_SiteFlagDoesNotMatchLabelVoltage_ReturnsEmpty()
        {
            _generator.Pat.SiteFlag = "HV@site-var=2;";
            _generator.LabelVoltage = "NV";

            string result = _generator.CreateIfConditionByTestPanDefine();

            Assert.AreEqual(string.Empty, result);
        }
        #endregion

        #region CreateShmooTName
        [TestMethod]
        public void CreateShmooTName_ReplacesVoltageTokenWithLabelVoltage()
        {
            _generator.Pat.Shmoo = new HardipCharSetup { TestNameInFlow = "SHMOO_TEST_V_1" };
            _generator.LabelVoltage = "NV";

            string result = _generator.CreateShmooTName();

            Assert.AreEqual("SHMOO_TEST_N_1", result);
        }

        [TestMethod]
        public void CreateShmooTName_StartsWithHac_ReturnsEmpty()
        {
            _generator.Pat.Shmoo = new HardipCharSetup { TestNameInFlow = "HAC_SHMOO_V_1" };
            _generator.LabelVoltage = "NV";

            string result = _generator.CreateShmooTName();

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void CreateShmooTName_NoShmoo_ReturnsEmpty()
        {
            _generator.Pat.Shmoo = null;

            string result = _generator.CreateShmooTName();

            Assert.AreEqual(string.Empty, result);
        }
        #endregion
    }
}
