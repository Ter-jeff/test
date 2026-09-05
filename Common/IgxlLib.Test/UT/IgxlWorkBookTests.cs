using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT
{
    [TestClass]
    public class IgxlWorkBookTests
    {
        private IgxlWorkBook _workBook;

        [TestInitialize]
        public void Setup()
        {
            _workBook = new IgxlWorkBook();
        }

        [TestMethod]
        public void IgxlWorkBook_Constructor_CreatesEmptyWorkBook()
        {
            Assert.AreEqual(0, _workBook.SubFlowSheets.Count);
            Assert.AreEqual(0, _workBook.MainFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddSubFlowSheet_SuccessfullyAddsSheet()
        {
            var subFlowSheet = new SubFlowSheet("TestFlow", "TestSource");
            _workBook.AddSubFlowSheet("", subFlowSheet);

            Assert.AreEqual(1, _workBook.SubFlowSheets.Count);
            Assert.IsTrue(_workBook.SubFlowSheets.ContainsKey("\\TestFlow"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddMainFlowSheet_SuccessfullyAddsSheet()
        {
            var mainFlow = new MainFlow("MainFlowTest", "TestSource");
            _workBook.AddMainFlowSheet("", mainFlow);

            Assert.AreEqual(1, _workBook.MainFlowSheets.Count);
            Assert.IsTrue(_workBook.MainFlowSheets.ContainsKey("\\MainFlowTest"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddInsSheet_SuccessfullyAddsInstanceSheet()
        {
            var instanceSheet = new InstanceSheet("TestInstance");
            _workBook.AddInsSheet("", instanceSheet);

            Assert.AreEqual(1, _workBook.InsSheets.Count);
            Assert.IsTrue(_workBook.InsSheets.ContainsKey("\\TestInstance"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddDcSpecSheet_SuccessfullyAddsSheet()
        {
            var dcSpecSheet = new DcSpecSheet("TestDcSpec");
            _workBook.AddDcSpecSheet("", dcSpecSheet);

            Assert.AreEqual(1, _workBook.DcSpecSheets.Count);
            Assert.IsTrue(_workBook.DcSpecSheets.ContainsKey("\\TestDcSpec"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddAcSpecSheet_SuccessfullyAddsSheet()
        {
            var acSpecSheet = new AcSpecSheet("TestAcSpec");
            _workBook.AddAcSpecSheet("", acSpecSheet);

            Assert.AreEqual(1, _workBook.AcSpecSheets.Count);
            Assert.IsTrue(_workBook.AcSpecSheets.ContainsKey("\\TestAcSpec"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddLevelSheet_SuccessfullyAddsSheet()
        {
            var levelSheet = new LevelSheet("TestLevel");
            _workBook.AddLevelSheet("", levelSheet);

            Assert.AreEqual(1, _workBook.LevelSheets.Count);
            Assert.IsTrue(_workBook.LevelSheets.ContainsKey("\\TestLevel"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddTimeSetSheet_SuccessfullyAddsSheet()
        {
            var timeSetSheet = new TimeSetBasicSheet("TestTimeSet");
            _workBook.AddTimeSetSheet("", timeSetSheet);

            Assert.AreEqual(1, _workBook.TimeSetSheets.Count);
            Assert.IsTrue(_workBook.TimeSetSheets.ContainsKey("\\TestTimeSet"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddPatSetSheet_SuccessfullyAddsSheet()
        {
            var patSetSheet = new PatSetSheet("TestPatSet");
            _workBook.AddPatSetSheet("", patSetSheet);

            Assert.AreEqual(1, _workBook.PatSetSheets.Count);
            Assert.IsTrue(_workBook.PatSetSheets.ContainsKey("\\TestPatSet"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddBinTblSheet_SuccessfullyAddsSheet()
        {
            var binTblSheet = new BinTableSheet("TestBinTable");
            _workBook.AddBinTblSheet("", binTblSheet);

            Assert.AreEqual(1, _workBook.BinTblSheets.Count);
            Assert.IsTrue(_workBook.BinTblSheets.ContainsKey("\\TestBinTable"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddBinTblSheet_MergesRowsForDuplicateSheet()
        {
            var binTblSheet1 = new BinTableSheet("BinTable");
            var row1 = new BinTableRow { Bin = "1", Result = "Pass" };
            binTblSheet1.Rows.Add(row1);

            var binTblSheet2 = new BinTableSheet("BinTable");
            var row2 = new BinTableRow { Bin = "2", Result = "Fail" };
            binTblSheet2.Rows.Add(row2);

            _workBook.AddBinTblSheet("", binTblSheet1);
            _workBook.AddBinTblSheet("", binTblSheet2);

            Assert.AreEqual(1, _workBook.BinTblSheets.Count);
            Assert.AreEqual(2, _workBook.BinTblSheets["\\BinTable"].Rows.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetFlowSheet_RetrievesExistingSheet()
        {
            var subFlowSheet = new SubFlowSheet("TestFlow", "TestSource");
            _workBook.AddSubFlowSheet("", subFlowSheet);

            SubFlowSheet retrievedSheet = _workBook.GetFlowSheet("TestFlow", "");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("TestFlow", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetFlowSheet_CreatesNewSheetIfNotExists()
        {
            SubFlowSheet newSheet = _workBook.GetFlowSheet("NewFlow", "");

            Assert.IsNotNull(newSheet);
            Assert.AreEqual("NewFlow", newSheet.Name);
            Assert.AreEqual(1, _workBook.SubFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPatSetsSheet_RetrievesExistingSheet()
        {
            var patSetSheet = new PatSetSheet("PatSet1");
            _workBook.AddPatSetSheet("", patSetSheet);

            PatSetSheet retrievedSheet = _workBook.GetPatSetsSheet("PatSet1", "");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("PatSet1", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPatSetSubSheet_RetrievesExistingSheet()
        {
            var patSetSubSheet = new PatSetSubSheet("PatSetSub1");
            _workBook.AddPatSetSubSheet("", patSetSubSheet);

            PatSetSubSheet retrievedSheet = _workBook.GetPatSetSubSheet("PatSetSub1", "");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("PatSetSub1", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetInstanceSheet_RetrievesExistingSheet()
        {
            var instanceSheet = new InstanceSheet("TestInst");
            _workBook.AddInsSheet("", instanceSheet);

            InstanceSheet retrievedSheet = _workBook.GetInstanceSheet("TestInst");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("TestInst", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetInstanceSheet_ReturnsNullForNonexistentSheet()
        {
            InstanceSheet retrievedSheet = _workBook.GetInstanceSheet("NonExistent");

            Assert.IsNull(retrievedSheet);
        }

        [TestMethod]
        public void IgxlWorkBook_GetAcSpecsSheet_RetrievesExistingSheet()
        {
            var acSpecSheet = new AcSpecSheet("AC_Specs");
            _workBook.AddAcSpecSheet("", acSpecSheet);

            AcSpecSheet retrievedSheet = _workBook.GetAcSpecsSheet("AC_Specs");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("AC_Specs", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetBinTblSheet_RetrievesExistingSheet()
        {
            var binTblSheet = new BinTableSheet("Bin_Table");
            _workBook.AddBinTblSheet("", binTblSheet);

            BinTableSheet retrievedSheet = _workBook.GetBinTblSheet("", "Bin_Table");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("Bin_Table", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetBinTblSheet_CreatesNewSheetIfNotExists()
        {
            BinTableSheet retrievedSheet = _workBook.GetBinTblSheet("", "NewBinTable");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("NewBinTable", retrievedSheet.Name);
            Assert.AreEqual(1, _workBook.BinTblSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetMainBinTblSheet_RetrievesMainBinTable()
        {
            var binTblSheet = new BinTableSheet("Bin_Table");
            _workBook.AddBinTblSheet("", binTblSheet);

            BinTableSheet mainBinSheet = _workBook.GetMainBinTblSheet("");

            Assert.IsNotNull(mainBinSheet);
            Assert.AreEqual("Bin_Table", mainBinSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_Contact_ConcatenatesPathsCorrectly()
        {
            string result = IgxlWorkBook.Contact("dir1", "dir2");

            Assert.AreEqual($"dir1{Path.DirectorySeparatorChar}dir2", result);
        }

        [TestMethod]
        public void IgxlWorkBook_Contact_HandlesTrailingAndLeadingSlashes()
        {
            string result = IgxlWorkBook.Contact("dir1\\", "/dir2");

            Assert.AreEqual($"dir1{Path.DirectorySeparatorChar}dir2", result);
        }

        [TestMethod]
        public void IgxlWorkBook_Contact_ThrowsExceptionForNullMainDir()
        {
            Assert.ThrowsException<Exception>(() => IgxlWorkBook.Contact(null, "dir2"));
        }

        [TestMethod]
        public void IgxlWorkBook_Contact_ThrowsExceptionForNullSubDir()
        {
            Assert.ThrowsException<Exception>(() => IgxlWorkBook.Contact("dir1", null));
        }

        [TestMethod]
        public void IgxlWorkBook_GetSubFileSubName_ReturnsExpectedPath()
        {
            string result = IgxlWorkBook.GetSubFileSubName("", "TestName");

            Assert.AreEqual("\\TestName", result);
        }

        [TestMethod]
        public void IgxlWorkBook_GetFlowSheetNameList_ReturnsAllFlowNames()
        {
            var flow1 = new SubFlowSheet("Flow1", "Source1");
            var flow2 = new SubFlowSheet("Flow2", "Source2");
            _workBook.AddSubFlowSheet("Dir1", flow1);
            _workBook.AddSubFlowSheet("Dir2", flow2);

            List<string> flowNames = _workBook.GetFlowSheetNameList();

            Assert.AreEqual(2, flowNames.Count);
            Assert.IsTrue(flowNames.Contains("Flow1"));
            Assert.IsTrue(flowNames.Contains("Flow2"));
        }

        [TestMethod]
        public void IgxlWorkBook_Clear_RemovesAllSheets()
        {
            _workBook.AddSubFlowSheet("Dir", new SubFlowSheet("Flow", "Source"));
            _workBook.AddInsSheet("Dir", new InstanceSheet("Instance"));
            _workBook.AddLevelSheet("Dir", new LevelSheet("Level"));

            Assert.IsTrue(_workBook.SubFlowSheets.Count > 0);
            Assert.IsTrue(_workBook.InsSheets.Count > 0);
            Assert.IsTrue(_workBook.LevelSheets.Count > 0);

            _workBook.Clear();

            Assert.AreEqual(0, _workBook.SubFlowSheets.Count);
            Assert.AreEqual(0, _workBook.InsSheets.Count);
            Assert.AreEqual(0, _workBook.LevelSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AllIgxlSheetsDicWithType_ContainsAllSheetTypes()
        {
            _workBook.AddSubFlowSheet("Dir", new SubFlowSheet("Flow", "Source"));
            _workBook.AddInsSheet("Dir", new InstanceSheet("Instance"));
            _workBook.AddLevelSheet("Dir", new LevelSheet("Level"));
            _workBook.AddDcSpecSheet("Dir", new DcSpecSheet("DcSpec"));

            Dictionary<string, string> allSheets = _workBook.AllIgxlSheetsDicWithType;

            Assert.IsTrue(allSheets.Count >= 4);
            Assert.IsTrue(allSheets.ContainsValue("Flow Table"));
            Assert.IsTrue(allSheets.ContainsValue("Test Instances"));
            Assert.IsTrue(allSheets.ContainsValue("Pin Levels"));
            Assert.IsTrue(allSheets.ContainsValue("DC Specs"));
        }

        [TestMethod]
        public void IgxlWorkBook_AllIgxlSheets_ReturnsAllAddedSheets()
        {
            _workBook.AddSubFlowSheet("Dir", new SubFlowSheet("Flow", "Source"));
            _workBook.AddInsSheet("Dir", new InstanceSheet("Instance"));
            _workBook.AddLevelSheet("Dir", new LevelSheet("Level"));

            Dictionary<string, IIgxlSheet> allSheets = _workBook.AllIgxlSheets;

            Assert.AreEqual(3, allSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_SetFlowSheet_UpdatesExistingSheet()
        {
            var originalFlow = new SubFlowSheet("TestFlow", "Original");
            _workBook.AddSubFlowSheet("Dir", originalFlow);

            var newFlow = new SubFlowSheet("TestFlow", "Updated");
            _workBook.SetFlowSheet("TestFlow", newFlow);

            SubFlowSheet retrievedFlow = _workBook.GetFlowSheet("TestFlow", "Dir");
            Assert.AreEqual("Updated", retrievedFlow.SourceInfo.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_TryGetTestInstCommon_ReturnsTrue_WhenInstanceExists()
        {
            var testInstCommon = new InstanceSheet("TestInst_Common");
            _workBook.AddInsSheet("Dir", testInstCommon);

            bool result = _workBook.TryGetTestInstCommon(out InstanceSheet instanceSheet);

            Assert.IsTrue(result);
            Assert.IsNotNull(instanceSheet);
            Assert.AreEqual("TestInst_Common", instanceSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_TryGetTestInstCommon_ReturnsFalse_WhenInstanceNotExists()
        {
            var otherInstance = new InstanceSheet("OtherInstance");
            _workBook.AddInsSheet("Dir", otherInstance);

            bool result = _workBook.TryGetTestInstCommon(out InstanceSheet instanceSheet);

            Assert.IsFalse(result);
            Assert.IsNull(instanceSheet);
        }

        [TestMethod]
        public void IgxlWorkBook_AddChannelMapSheet_SuccessfullyAddsSheet()
        {
            var channelMapSheet = new ChannelMapSheet("TestChannelMap");
            _workBook.AddChannelMapSheet("", channelMapSheet);

            Assert.AreEqual(1, _workBook.ChannelMapSheets.Count);
            Assert.IsTrue(_workBook.ChannelMapSheets.ContainsKey("\\TestChannelMap"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddPortMapSheet_SuccessfullyAddsSheet()
        {
            var portMapSheet = new PortMapSheet("TestPortMap");
            _workBook.AddPortMapSheet("", portMapSheet);

            Assert.AreEqual(1, _workBook.PortMapSheets.Count);
            Assert.IsTrue(_workBook.PortMapSheets.ContainsKey("\\TestPortMap"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddWaveDefSheet_SuccessfullyAddsSheet()
        {
            var waveDefSheet = new WaveDefinitionSheet("TestWaveDef");
            _workBook.AddWaveDefSheet("", waveDefSheet);

            Assert.AreEqual(1, _workBook.WaveDefSheets.Count);
            Assert.IsTrue(_workBook.WaveDefSheets.ContainsKey("\\TestWaveDef"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddMixedSignalSheet_SuccessfullyAddsSheet()
        {
            var mixedSignalSheet = new MixedSignalSheet("TestMixedSignal");
            _workBook.AddMixedSignalSheet("", mixedSignalSheet);

            Assert.AreEqual(1, _workBook.MixedSignalSheets.Count);
            Assert.IsTrue(_workBook.MixedSignalSheets.ContainsKey("\\TestMixedSignal"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddCharSheet_SuccessfullyAddsSheet()
        {
            var charSheet = new CharSheet("TestChar");
            _workBook.AddCharSheet("", charSheet);

            Assert.AreEqual(1, _workBook.CharSheets.Count);
            Assert.IsTrue(_workBook.CharSheets.ContainsKey("\\TestChar"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddJitterSheet_SuccessfullyAddsSheet()
        {
            var jitterSheet = new JitterSheet("TestJitter");
            _workBook.AddJitterSheet("", jitterSheet);

            Assert.AreEqual(1, _workBook.JitterSheets.Count);
            Assert.IsTrue(_workBook.JitterSheets.ContainsKey("\\TestJitter"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddReferenceSheet_SuccessfullyAddsSheet()
        {
            var referenceSheet = new ReferenceSheet("TestReference");
            _workBook.AddReferenceSheet("", referenceSheet);

            Assert.AreEqual(1, _workBook.ReferenceSheets.Count);
            Assert.IsTrue(_workBook.ReferenceSheets.ContainsKey("\\TestReference"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddJobListSheet_SuccessfullyAddsSheet()
        {
            var jobListSheet = new JobListSheet("TestJobList");
            _workBook.AddJobListSheet("", jobListSheet);

            Assert.AreEqual(1, _workBook.JobListSheets.Count);
            Assert.IsTrue(_workBook.JobListSheets.ContainsKey("\\TestJobList"));
        }

        [TestMethod]
        public void IgxlWorkBook_SetGlbSpecSheetPair_SuccessfullySetsPair()
        {
            var globalSpecSheet = new GlobalSpecSheet("TestGlobalSpec");
            var pair = new KeyValuePair<string, GlobalSpecSheet>("", globalSpecSheet);

            _workBook.GlbSpecSheetPair = pair;

            Assert.AreEqual("\\TestGlobalSpec", _workBook.GlbSpecSheetPair.Key);
            Assert.AreEqual(globalSpecSheet, _workBook.GlbSpecSheetPair.Value);
        }

        [TestMethod]
        public void IgxlWorkBook_SetGlbSpecSheetPair_ThrowsExceptionWhenAlreadySet()
        {
            var globalSpecSheet1 = new GlobalSpecSheet("TestGlobalSpec1");
            var pair1 = new KeyValuePair<string, GlobalSpecSheet>("TestKey1", globalSpecSheet1);
            _workBook.GlbSpecSheetPair = pair1;

            var globalSpecSheet2 = new GlobalSpecSheet("TestGlobalSpec2");
            var pair2 = new KeyValuePair<string, GlobalSpecSheet>("TestKey2", globalSpecSheet2);

            Assert.ThrowsException<Exception>(() => _workBook.GlbSpecSheetPair = pair2);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPinMapPair_ReturnsDefaultPairInitially()
        {
            KeyValuePair<string, PinMapSheet> pair = _workBook.PinMapPair;

            Assert.IsNotNull(pair.Value);
            Assert.AreEqual("", pair.Value.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_SetPinMapPair_SuccessfullySetsPair()
        {
            var pinMapSheet = new PinMapSheet("TestPinMap");
            var pair = new KeyValuePair<string, PinMapSheet>("", pinMapSheet);

            _workBook.PinMapPair = pair;

            Assert.AreEqual("\\TestPinMap", _workBook.PinMapPair.Key);
            Assert.AreEqual(pinMapSheet, _workBook.PinMapPair.Value);
        }

        [TestMethod]
        public void IgxlWorkBook_SetPinMapPair_UpdatesExistingPairIfEmpty()
        {
            var pinMapSheet1 = new PinMapSheet("");
            var pair1 = new KeyValuePair<string, PinMapSheet>("", pinMapSheet1);
            _workBook.PinMapPair = pair1;

            var pinMapSheet2 = new PinMapSheet("TestPinMap");
            var pair2 = new KeyValuePair<string, PinMapSheet>("", pinMapSheet2);
            _workBook.PinMapPair = pair2;

            Assert.AreEqual("\\TestPinMap", _workBook.PinMapPair.Key);
            Assert.AreEqual(pinMapSheet2, _workBook.PinMapPair.Value);
        }

        [TestMethod]
        public void IgxlWorkBook_GetRestrictedIgxlSheets_ReturnsEmptyDictionary()
        {
            Dictionary<string, IIgxlSheet> restrictedSheets = _workBook.RestrictedIgxlSheets;

            Assert.IsNotNull(restrictedSheets);
            Assert.AreEqual(0, restrictedSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GenEnableToMainInitEnableWd_SuccessfullyAddsFlowRow()
        {
            var subFlowSheet = new SubFlowSheet("Flow_Table_Main_Init_EnableWd", "Initialize");
            _workBook.AddSubFlowSheet("TestDir", subFlowSheet);

            _workBook.GenEnableToMainInitEnableWd("enable_flag", "FlowOpcode", "TestColumnA", "TestDir");

            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual("enable_flag", subFlowSheet.Rows[0].Enable);
            Assert.AreEqual("FlowOpcode", subFlowSheet.Rows[0].Opcode);
            Assert.AreEqual("TestColumnA", subFlowSheet.Rows[0].ColumnA);
        }

        [TestMethod]
        public void IgxlWorkBook_GenEnableToMainInitEnableWd_WithParameter_SuccessfullyAddsFlowRow()
        {
            var subFlowSheet = new SubFlowSheet("Flow_Table_Main_Init_EnableWd", "Initialize");
            _workBook.AddSubFlowSheet("TestDir", subFlowSheet);

            _workBook.GenEnableToMainInitEnableWd("enable_flag", "FlowOpcode", "TestColumnA", "TestDir", "TestParameter");

            Assert.AreEqual(1, subFlowSheet.Rows.Count);
            Assert.AreEqual("TestParameter", subFlowSheet.Rows[0].Parameter);
        }

        [TestMethod]
        public void IgxlWorkBook_GenEnableToMainInitEnableWd_WithoutFlowSheet_DoesNotThrow()
        {
            // This method should not throw an exception if the flow sheet doesn't exist
            _workBook.GenEnableToMainInitEnableWd("enable_flag", "FlowOpcode", "TestColumnA", "NonExistentDir");

            Assert.AreEqual(1, _workBook.SubFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPatSetsSheet_CreatesNewSheetIfNotExists()
        {
            PatSetSheet patSetSheet = _workBook.GetPatSetsSheet("NewPatSet", "TestDir");

            Assert.IsNotNull(patSetSheet);
            Assert.AreEqual("NewPatSet", patSetSheet.Name);
            Assert.AreEqual(1, _workBook.PatSetSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPatSetSubSheet_CreatesNewSheetIfNotExists()
        {
            PatSetSubSheet patSetSubSheet = _workBook.GetPatSetSubSheet("NewPatSetSub", "TestDir");

            Assert.IsNotNull(patSetSubSheet);
            Assert.AreEqual("NewPatSetSub", patSetSubSheet.Name);
            Assert.AreEqual(1, _workBook.PatSetSubSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetAcSpecsSheet_WithDefaultName_RetrievesSheet()
        {
            var acSpecSheet = new AcSpecSheet("AC_Specs");
            _workBook.AddAcSpecSheet("TestDir", acSpecSheet);

            AcSpecSheet retrievedSheet = _workBook.GetAcSpecsSheet();

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("AC_Specs", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetAcSpecsSheet_WithCustomName_RetrievesSheet()
        {
            var acSpecSheet = new AcSpecSheet("CustomAcSpecs");
            _workBook.AddAcSpecSheet("TestDir", acSpecSheet);

            AcSpecSheet retrievedSheet = _workBook.GetAcSpecsSheet("CustomAcSpecs");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("CustomAcSpecs", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetInstanceSheet_CaseInsensitive()
        {
            var instanceSheet = new InstanceSheet("TestInstance");
            _workBook.AddInsSheet("TestDir", instanceSheet);

            InstanceSheet retrievedSheet = _workBook.GetInstanceSheet("testinstance");

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("TestInstance", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_AddSubFlowSheet_KeyValuePair_SuccessfullyAddsSheet()
        {
            var subFlowSheet = new SubFlowSheet("TestFlow", "TestSource");
            var pair = new KeyValuePair<string, SubFlowSheet>("TestKey", subFlowSheet);

            _workBook.AddSubFlowSheet(pair);

            Assert.AreEqual(1, _workBook.SubFlowSheets.Count);
            Assert.IsTrue(_workBook.SubFlowSheets.ContainsKey("TestKey"));
        }

        [TestMethod]
        public void IgxlWorkBook_GenAllFailFlagToMainInitEnableWd_WithoutFlowSheet_DoesNotThrow()
        {
            var flags = new List<string> { "flag1", "flag2" };

            // Should not throw even if flow sheet doesn't exist
            _workBook.GenAllFailFlagToMainInitEnableWd(flags, "TestColumnA", "NonExistentDir");

            Assert.AreEqual(1, _workBook.SubFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GenAllFailFlagToMainInitEnableWd_WithEmptyFlags_DoesNothing()
        {
            var subFlowSheet = new SubFlowSheet("Flow_Table_Main_Init_EnableWd", "Initialize");
            _workBook.AddSubFlowSheet("TestDir", subFlowSheet);

            var emptyFlags = new List<string>();
            _workBook.GenAllFailFlagToMainInitEnableWd(emptyFlags, "TestColumnA", "TestDir");

            Assert.AreEqual(0, subFlowSheet.Rows.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_SetFlowSheet_NonexistentSheet_DoesNotUpdate()
        {
            var newFlow = new SubFlowSheet("NonExistentFlow", "NewSource");
            _workBook.SetFlowSheet("NonExistentFlow", newFlow);

            Assert.AreEqual(0, _workBook.SubFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_LimitSetsSheets_PropertyReturnsCollection()
        {
            Assert.AreEqual(0, _workBook.LimitSetsSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetFlowSheet_WithSourceParameter_SuccessfullyCreatesSheet()
        {
            SubFlowSheet flowSheet = _workBook.GetFlowSheet("TestFlow", "TestDir", "CustomSource");

            Assert.IsNotNull(flowSheet);
            Assert.AreEqual("TestFlow", flowSheet.Name);
            Assert.AreEqual("CustomSource", flowSheet.SourceInfo.Name);
            Assert.AreEqual(1, _workBook.SubFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetAcSpecsSheet_WithDefaultParameter_RetrievesAcSpecs()
        {
            var acSpecSheet = new AcSpecSheet("AC_Specs");
            _workBook.AddAcSpecSheet("", acSpecSheet);

            AcSpecSheet retrievedSheet = _workBook.GetAcSpecsSheet();

            Assert.IsNotNull(retrievedSheet);
            Assert.AreEqual("AC_Specs", retrievedSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_AddDcSpecSheet_MultipleSheets_AllAreAdded()
        {
            var dcSpec1 = new DcSpecSheet("DcSpec1");
            var dcSpec2 = new DcSpecSheet("DcSpec2");
            var dcSpec3 = new DcSpecSheet("DcSpec3");

            _workBook.AddDcSpecSheet("Dir1", dcSpec1);
            _workBook.AddDcSpecSheet("Dir2", dcSpec2);
            _workBook.AddDcSpecSheet("Dir3", dcSpec3);

            Assert.AreEqual(3, _workBook.DcSpecSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddLevelSheet_MultipleSheets_AllAreAdded()
        {
            var level1 = new LevelSheet("Level1");
            var level2 = new LevelSheet("Level2");
            var level3 = new LevelSheet("Level3");

            _workBook.AddLevelSheet("Dir1", level1);
            _workBook.AddLevelSheet("Dir2", level2);
            _workBook.AddLevelSheet("Dir3", level3);

            Assert.AreEqual(3, _workBook.LevelSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddTimeSetSheet_MultipleSheets_AllAreAdded()
        {
            var timeSet1 = new TimeSetBasicSheet("TimeSet1");
            var timeSet2 = new TimeSetBasicSheet("TimeSet2");

            _workBook.AddTimeSetSheet("Dir1", timeSet1);
            _workBook.AddTimeSetSheet("Dir2", timeSet2);

            Assert.AreEqual(2, _workBook.TimeSetSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddPatSetSubSheet_MultipleSheets_AllAreAdded()
        {
            var patSetSub1 = new PatSetSubSheet("PatSetSub1");
            var patSetSub2 = new PatSetSubSheet("PatSetSub2");

            _workBook.AddPatSetSubSheet("Dir1", patSetSub1);
            _workBook.AddPatSetSubSheet("Dir2", patSetSub2);

            Assert.AreEqual(2, _workBook.PatSetSubSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddMainFlowSheet_MultipleFlows_AllAreAdded()
        {
            var mainFlow1 = new MainFlow("MainFlow1", "Source1");
            var mainFlow2 = new MainFlow("MainFlow2", "Source2");
            var mainFlow3 = new MainFlow("MainFlow3", "Source3");

            _workBook.AddMainFlowSheet("Dir1", mainFlow1);
            _workBook.AddMainFlowSheet("Dir2", mainFlow2);
            _workBook.AddMainFlowSheet("Dir3", mainFlow3);

            Assert.AreEqual(3, _workBook.MainFlowSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddInstanceSheet_MultipleSheets_AllAreAdded()
        {
            var inst1 = new InstanceSheet("Instance1");
            var inst2 = new InstanceSheet("Instance2");
            var inst3 = new InstanceSheet("Instance3");

            _workBook.AddInsSheet("Dir1", inst1);
            _workBook.AddInsSheet("Dir2", inst2);
            _workBook.AddInsSheet("Dir3", inst3);

            Assert.AreEqual(3, _workBook.InsSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_Clear_ClearsAllSheetTypes()
        {
            _workBook.AddSubFlowSheet("Dir", new SubFlowSheet("Flow", "Source"));
            _workBook.AddMainFlowSheet("Dir", new MainFlow("MainFlow", "Source"));
            _workBook.AddInsSheet("Dir", new InstanceSheet("Instance"));
            _workBook.AddDcSpecSheet("Dir", new DcSpecSheet("DcSpec"));
            _workBook.AddAcSpecSheet("Dir", new AcSpecSheet("AcSpec"));
            _workBook.AddLevelSheet("Dir", new LevelSheet("Level"));
            _workBook.AddTimeSetSheet("Dir", new TimeSetBasicSheet("TimeSet"));
            _workBook.AddPatSetSheet("Dir", new PatSetSheet("PatSet"));
            _workBook.AddPatSetSubSheet("Dir", new PatSetSubSheet("PatSetSub"));
            _workBook.AddBinTblSheet("Dir", new BinTableSheet("BinTable"));
            _workBook.AddChannelMapSheet("Dir", new ChannelMapSheet("ChannelMap"));
            _workBook.AddJobListSheet("Dir", new JobListSheet("JobList"));
            _workBook.AddPortMapSheet("Dir", new PortMapSheet("PortMap"));
            _workBook.AddWaveDefSheet("Dir", new WaveDefinitionSheet("WaveDef"));
            _workBook.AddMixedSignalSheet("Dir", new MixedSignalSheet("MixedSignal"));
            _workBook.AddCharSheet("Dir", new CharSheet("Char"));
            _workBook.AddJitterSheet("Dir", new JitterSheet("Jitter"));
            _workBook.AddReferenceSheet("Dir", new ReferenceSheet("Reference"));

            int totalBefore = _workBook.AllIgxlSheets.Count;
            Assert.IsTrue(totalBefore > 0);

            _workBook.Clear();

            Assert.AreEqual(0, _workBook.SubFlowSheets.Count);
            Assert.AreEqual(0, _workBook.MainFlowSheets.Count);
            Assert.AreEqual(0, _workBook.InsSheets.Count);
            Assert.AreEqual(0, _workBook.DcSpecSheets.Count);
            Assert.AreEqual(0, _workBook.AcSpecSheets.Count);
            Assert.AreEqual(0, _workBook.LevelSheets.Count);
            Assert.AreEqual(0, _workBook.TimeSetSheets.Count);
            Assert.AreEqual(0, _workBook.PatSetSheets.Count);
            Assert.AreEqual(0, _workBook.PatSetSubSheets.Count);
            Assert.AreEqual(0, _workBook.BinTblSheets.Count);
            Assert.AreEqual(0, _workBook.ChannelMapSheets.Count);
            Assert.AreEqual(0, _workBook.JobListSheets.Count);
            Assert.AreEqual(0, _workBook.PortMapSheets.Count);
            Assert.AreEqual(0, _workBook.WaveDefSheets.Count);
            Assert.AreEqual(0, _workBook.MixedSignalSheets.Count);
            Assert.AreEqual(0, _workBook.CharSheets.Count);
            Assert.AreEqual(0, _workBook.JitterSheets.Count);
            Assert.AreEqual(0, _workBook.ReferenceSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_Contact_EmptyMainDir_ConcatenatesCorrectly()
        {
            string result = IgxlWorkBook.Contact("", "dir2");

            Assert.AreEqual($"{Path.DirectorySeparatorChar}dir2", result);
        }

        [TestMethod]
        public void IgxlWorkBook_GetFlowSheetNameList_WithMultipleFlows_ReturnsAllNames()
        {
            var flow1 = new SubFlowSheet("Flow1", "Source1");
            var flow2 = new SubFlowSheet("Flow2", "Source2");
            var flow3 = new SubFlowSheet("Flow3", "Source3");

            _workBook.AddSubFlowSheet("Dir1", flow1);
            _workBook.AddSubFlowSheet("Dir2", flow2);
            _workBook.AddSubFlowSheet("Dir3", flow3);

            List<string> flowNames = _workBook.GetFlowSheetNameList();

            Assert.AreEqual(3, flowNames.Count);
            Assert.IsTrue(flowNames.Contains("Flow1"));
            Assert.IsTrue(flowNames.Contains("Flow2"));
            Assert.IsTrue(flowNames.Contains("Flow3"));
        }

        [TestMethod]
        public void IgxlWorkBook_GetSubFileSubName_WithNonEmptyPath_ReturnsContactPath()
        {
            string result = IgxlWorkBook.GetSubFileSubName("Dir1/Dir2", "SheetName");

            string expected = $"Dir1{Path.DirectorySeparatorChar}Dir2{Path.DirectorySeparatorChar}SheetName";
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void IgxlWorkBook_AllIgxlSheetsDicWithType_IncludesAllSheetTypesWithCorrectNames()
        {
            _workBook.AddSubFlowSheet("Dir", new SubFlowSheet("SubFlow", "Source"));
            _workBook.AddMainFlowSheet("Dir", new MainFlow("MainFlow", "Source"));
            _workBook.AddInsSheet("Dir", new InstanceSheet("Instance"));
            _workBook.AddDcSpecSheet("Dir", new DcSpecSheet("DcSpec"));

            Dictionary<string, string> allSheets = _workBook.AllIgxlSheetsDicWithType;

            Assert.IsTrue(allSheets.Count >= 4);
            // MainFlow should be identified as "Main Flow"
            bool hasMainFlow = allSheets.Values.Any(v => v == "Main Flow");
            Assert.IsTrue(hasMainFlow);
        }

        [TestMethod]
        public void IgxlWorkBook_GetFlowSheet_ExistingSheet_ReturnsExistingInstance()
        {
            var originalFlow = new SubFlowSheet("TestFlow", "Original");
            _workBook.AddSubFlowSheet("Dir", originalFlow);

            // Get the same flow - should return the same instance
            SubFlowSheet retrievedFlow = _workBook.GetFlowSheet("TestFlow", "Dir");

            Assert.AreEqual("Original", retrievedFlow.SourceInfo.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPatSetsSheet_ExistingSheet_ReturnsSameInstance()
        {
            var originalPatSet = new PatSetSheet("PatSet1");
            _workBook.AddPatSetSheet("Dir", originalPatSet);

            PatSetSheet retrievedPatSet = _workBook.GetPatSetsSheet("PatSet1", "Dir");

            Assert.AreEqual(originalPatSet, retrievedPatSet);
        }

        [TestMethod]
        public void IgxlWorkBook_GetPatSetSubSheet_ExistingSheet_ReturnsSameInstance()
        {
            var originalPatSetSub = new PatSetSubSheet("PatSetSub1");
            _workBook.AddPatSetSubSheet("Dir", originalPatSetSub);

            PatSetSubSheet retrievedPatSetSub = _workBook.GetPatSetSubSheet("PatSetSub1", "Dir");

            Assert.AreEqual(originalPatSetSub, retrievedPatSetSub);
        }

        [TestMethod]
        public void IgxlWorkBook_TryGetTestInstCommon_CaseInsensitive()
        {
            var testInstCommon = new InstanceSheet("testinst_common");
            _workBook.AddInsSheet("Dir", testInstCommon);

            bool result = _workBook.TryGetTestInstCommon(out InstanceSheet instanceSheet);

            Assert.IsTrue(result);
            Assert.IsNotNull(instanceSheet);
            Assert.AreEqual("testinst_common", instanceSheet.Name);
        }

        [TestMethod]
        public void IgxlWorkBook_AddBinTblSheet_WithPath_ContainsPathInKey()
        {
            var binTblSheet = new BinTableSheet("TestBinTable");
            _workBook.AddBinTblSheet("MyPath", binTblSheet);

            Assert.AreEqual(1, _workBook.BinTblSheets.Count);
            string key = _workBook.BinTblSheets.Keys.First();
            Assert.IsTrue(key.Contains("MyPath") || key.Contains("TestBinTable"));
        }

        [TestMethod]
        public void IgxlWorkBook_AddChannelMapSheet_WithPath_SuccessfullyAdds()
        {
            var channelMapSheet = new ChannelMapSheet("TestChannelMap");
            _workBook.AddChannelMapSheet("TestPath", channelMapSheet);

            Assert.AreEqual(1, _workBook.ChannelMapSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddPortMapSheet_WithPath_SuccessfullyAdds()
        {
            var portMapSheet = new PortMapSheet("TestPortMap");
            _workBook.AddPortMapSheet("TestPath", portMapSheet);

            Assert.AreEqual(1, _workBook.PortMapSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddWaveDefSheet_WithPath_SuccessfullyAdds()
        {
            var waveDefSheet = new WaveDefinitionSheet("TestWaveDef");
            _workBook.AddWaveDefSheet("TestPath", waveDefSheet);

            Assert.AreEqual(1, _workBook.WaveDefSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddMixedSignalSheet_WithPath_SuccessfullyAdds()
        {
            var mixedSignalSheet = new MixedSignalSheet("TestMixedSignal");
            _workBook.AddMixedSignalSheet("TestPath", mixedSignalSheet);

            Assert.AreEqual(1, _workBook.MixedSignalSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddCharSheet_WithPath_SuccessfullyAdds()
        {
            var charSheet = new CharSheet("TestChar");
            _workBook.AddCharSheet("TestPath", charSheet);

            Assert.AreEqual(1, _workBook.CharSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddJitterSheet_WithPath_SuccessfullyAdds()
        {
            var jitterSheet = new JitterSheet("TestJitter");
            _workBook.AddJitterSheet("TestPath", jitterSheet);

            Assert.AreEqual(1, _workBook.JitterSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddReferenceSheet_WithPath_SuccessfullyAdds()
        {
            var referenceSheet = new ReferenceSheet("TestReference");
            _workBook.AddReferenceSheet("TestPath", referenceSheet);

            Assert.AreEqual(1, _workBook.ReferenceSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_AddJobListSheet_WithPath_SuccessfullyAdds()
        {
            var jobListSheet = new JobListSheet("TestJobList");
            _workBook.AddJobListSheet("TestPath", jobListSheet);

            Assert.AreEqual(1, _workBook.JobListSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_MultipleAddOperations_AllSheetsPresent()
        {
            var flow = new SubFlowSheet("Flow", "Source");
            var instance = new InstanceSheet("Instance");
            var dcSpec = new DcSpecSheet("DcSpec");
            var level = new LevelSheet("Level");

            _workBook.AddSubFlowSheet("Dir1", flow);
            _workBook.AddInsSheet("Dir2", instance);
            _workBook.AddDcSpecSheet("Dir3", dcSpec);
            _workBook.AddLevelSheet("Dir4", level);

            Dictionary<string, IIgxlSheet> allSheets = _workBook.AllIgxlSheets;
            Assert.AreEqual(4, allSheets.Count);
        }

        [TestMethod]
        public void IgxlWorkBook_GetInstanceSheet_EmptyWorkbook_ReturnsNull()
        {
            InstanceSheet result = _workBook.GetInstanceSheet("NonExistent");

            Assert.IsNull(result);
        }

        [TestMethod]
        public void IgxlWorkBook_GetAcSpecsSheet_WithNonDefaultName_ReturnsSheet()
        {
            var customAcSpec = new AcSpecSheet("CustomAcSpecs");
            _workBook.AddAcSpecSheet("Dir", customAcSpec);

            AcSpecSheet result = _workBook.GetAcSpecsSheet("CustomAcSpecs");

            Assert.IsNotNull(result);
            Assert.AreEqual("CustomAcSpecs", result.Name);
        }

        [TestMethod]
        public void GetSubFlows_Should_Return_All_Rows()
        {
            // Arrange
            var entry = new FlowRow
            {
                Opcode = OpCode.Call,
                Parameter = "SheetA"
            };

            _workBook.SubFlowSheets["SheetA"] = new SubFlowSheet("SheetA")
            {
                Rows =
                    [
                        new() { Opcode = "Op1" },
                        new() { Opcode = "Op2" }
                    ]
            };

            // Act
            List<FlowRow> result = _workBook.GetSubFlows(entry);

            // Assert
            Assert.AreEqual(3, result.Count);
        }

        [TestMethod]
        public void GetSubFlows_Should_Handle_Nested_Call()
        {
            // Arrange
            var entry = new FlowRow
            {
                Opcode = OpCode.Call,
                Parameter = "SheetA"
            };

            _workBook.SubFlowSheets["SheetA"] = new SubFlowSheet("SheetA")
            {
                Rows =
                    [
                        new() { Opcode = OpCode.Call, Parameter = "SheetB" }
                    ]
            };
            _workBook.SubFlowSheets["SheetB"] = new SubFlowSheet("SheetB")
            {
                Rows =
                    [
                        new() { Opcode = "OpB" }
                    ]
            };

            // Act
            List<FlowRow> result = _workBook.GetSubFlows(entry);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Any(x => x.Opcode == "OpB"));
        }

        [TestMethod]
        public void GetSubFlows_Should_Filter_Opcode()
        {
            // Arrange
            var entry = new FlowRow
            {
                Opcode = OpCode.Call,
                Parameter = "SheetA"
            };

            _workBook.SubFlowSheets["SheetA"] = new SubFlowSheet("SheetA")
            {
                Rows =
                    [
                        new() { Opcode = "Keep" },
                        new() { Opcode = "Skip" }
                    ]
            };

            // Act
            List<FlowRow> result = _workBook.GetSubFlows(entry, "Keep");

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Any(x => x.Opcode == "Keep"));
            Assert.IsFalse(result.Any(x => x.Opcode == "Skip"));
        }

        [TestMethod]
        public void GetSubFlows_Should_Avoid_Cycle()
        {
            // Arrange
            var entry = new FlowRow
            {
                Opcode = OpCode.Call,
                Parameter = "SheetA"
            };

            _workBook.SubFlowSheets["SheetA"] = new SubFlowSheet("SheetA")
            {
                Rows =
                    [
                        new() { Opcode = OpCode.Call, Parameter = "SheetB" }
                    ]
            };
            _workBook.SubFlowSheets["SheetB"] = new SubFlowSheet("SheetB")
            {
                Rows =
                    [
                        new() { Opcode = OpCode.Call, Parameter = "SheetA" } // cycle
                    ]
            };

            // Act
            List<FlowRow> result = _workBook.GetSubFlows(entry);

            // Assert
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void GetSubFlows_Should_Skip_Backup_Rows()
        {
            // Arrange
            var entry = new FlowRow
            {
                Opcode = OpCode.Call,
                Parameter = "SheetA"
            };

            _workBook.SubFlowSheets["SheetA"] = new SubFlowSheet("SheetA")
            {
                Rows =
                    [
                        new() { Opcode = "Op1", IsBackup = true },
                        new() { Opcode = "Op2" }
                    ]
            };

            // Act
            List<FlowRow> result = _workBook.GetSubFlows(entry);

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsFalse(result.Any(x => x.IsBackup));
        }

    }
}
