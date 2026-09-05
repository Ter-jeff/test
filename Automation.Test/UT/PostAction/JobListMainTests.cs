using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenJob;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Basic;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class JobListMainTests
    {
        private string _root = null!;

        [TestInitialize]
        public void Setup()
        {
            // Reset static singletons.
            TestProgram.Clear();
            TestPlanStatic.Clear();
            LocalSpecs.Clear();

            // Do NOT call FolderStructure.Clear(). It will null out DirMain, DirJob, etc.

            // Prepare a valid target folder so FolderStructure.Dir* returns valid paths.
            _root = Path.Combine(Path.GetTempPath(), "JobListMain_UT_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            LocalSpecs.TarFolder = _root;

            // Fresh options for every test.
            LocalSpecs.Options = new Options
            {
                GenerateT0TXTestprogram = false
            };

            // Fresh JobMap.
            LocalSpecs.JobMap = [];
        }

        [TestCleanup]
        public void Cleanup()
        {
            try
            {
                if (!string.IsNullOrEmpty(_root) && Directory.Exists(_root))
                {
                    Directory.Delete(_root, true);
                }
            }
            catch
            {
                // Best effort cleanup only.
            }

            TestProgram.Clear();
            TestPlanStatic.Clear();
            LocalSpecs.Clear();
            // Do not call FolderStructure.Clear() here either.
        }

        // ------------------------------------------------------------
        // 01. Early return when JobMap empty
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Return_When_JobMap_Empty()
        {
            // Arrange
            LocalSpecs.JobMap.Clear();
            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            Assert.AreEqual(0, TestProgram.IgxlWorkBk.JobListSheets.Count);
            Assert.AreEqual(0, TestProgram.SubProgIgxlWorkBk.JobListSheets.Count);
            Assert.AreEqual(0, TestProgram.T0TxIgxlWorkBk.JobListSheets.Count);
        }

        // ------------------------------------------------------------
        // 02. Main JobList flow match, no concurrent
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Create_MainJobList_With_Flow_Match()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT1" } }
            };

            MainFlow mainFlow = new MainFlow("Main_Flow_FT1") { JobNames = [] };
            mainFlow.JobNames = ["FT1"];
            mainFlow.HasConcurrent = false;
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, mainFlow);

            DcSpecSheet dc = new DcSpecSheet("DC_Common", [], null!);
            TestProgram.IgxlWorkBk.AddDcSpecSheet(FolderStructure.DirDcSpec, dc);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            Assert.AreEqual(1, TestProgram.IgxlWorkBk.JobListSheets.Count);
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("FT1");

            Assert.AreNotEqual(null, row);
            Assert.AreEqual("Main_Flow_FT1", row!.FlowTable);
            Assert.AreEqual("DC_Common", row.DcSpecs);
            Assert.AreEqual(string.Empty, row.ConcurrentSequence);
        }

        // ------------------------------------------------------------
        // 03. MainFlow fallback to All, with concurrent sequence
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Use_Flow_With_All_And_Set_Concurrent()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "CP", new List<string> { "CP1" } }
            };

            MainFlow allFlow = new MainFlow("Main_Flow_AllCases") { JobNames = [] };
            allFlow.JobNames = ["All"];
            allFlow.HasConcurrent = true;
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, allFlow);

            DcSpecSheet dc = new DcSpecSheet("DC_Default", [], null!);
            TestProgram.IgxlWorkBk.AddDcSpecSheet(FolderStructure.DirDcSpec, dc);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("CP1");

            Assert.AreNotEqual(null, row);
            Assert.AreEqual("Main_Flow_AllCases", row!.FlowTable);
            Assert.AreEqual("", row.ConcurrentSequence);
        }

        // ------------------------------------------------------------
        // 04. Multiple DC Specs join when names contain job
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Join_Matching_DcSpecs()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT2" } }
            };

            MainFlow flow = new MainFlow("Main_Flow_FT2") { JobNames = [] };
            flow.JobNames = ["FT2"];
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);

            DcSpecSheet dc1 = new DcSpecSheet("DC_FT2_A", [], null!);
            DcSpecSheet dc2 = new DcSpecSheet("DC_B_FT2", [], null!);
            TestProgram.IgxlWorkBk.AddDcSpecSheet(FolderStructure.DirDcSpec, dc1);
            TestProgram.IgxlWorkBk.AddDcSpecSheet(FolderStructure.DirDcSpec, dc2);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("FT2");
            Assert.AreNotEqual(null, row);
            Assert.AreEqual("DC_B_FT2", row!.DcSpecs);
        }

        // ------------------------------------------------------------
        // 05. DC spec is empty when no DC sheets exist
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Set_Empty_DcSpec_When_No_DcSheets()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT3" } }
            };

            MainFlow flow = new MainFlow("Main_Flow_FT3") { JobNames = [] };
            flow.JobNames = ["FT3"];
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("FT3");
            Assert.AreNotEqual(null, row);
            Assert.AreEqual(string.Empty, row!.DcSpecs);
        }

        // ------------------------------------------------------------
        // 06. SubProgram JobList
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Create_SubProgram_JobList()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT1" } }
            };

            MainFlow mainFlow = new MainFlow("Main_Flow_FT1") { JobNames = [] };
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, mainFlow);

            MainFlow subFlow = new MainFlow("Main_Flow_Sub_Whatever") { JobNames = [] };
            TestProgram.SubProgIgxlWorkBk.AddMainFlowSheet(FolderStructure.DirSubProgram, subFlow);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            Assert.AreEqual(1, TestProgram.SubProgIgxlWorkBk.JobListSheets.Count);
            JobListSheet sheet = TestProgram.SubProgIgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("FT1");
            Assert.AreNotEqual(null, row);
            Assert.AreEqual(TestProgram.SubProgIgxlWorkBk.MainFlowSheets.First().Value.Name, row!.FlowTable);
        }

        // ------------------------------------------------------------
        // 07. T0TX JobList
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Create_T0Tx_JobList()
        {
            // Arrange
            LocalSpecs.Options.GenerateT0TXTestprogram = true;
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT1", "FT2" } }
            };

            // Provide at least one normal main flow so main job list exists.
            MainFlow mainFlow = new MainFlow("Main_Flow_FT1") { JobNames = [] };
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, mainFlow);

            // Provide T0TX flows that contain Room and Hot.
            MainFlow roomFlow = new MainFlow("T0TX_Room_M1") { JobNames = [] };
            MainFlow hotFlow = new MainFlow("T0TX_Hot_M2") { JobNames = [] };
            TestProgram.T0TxIgxlWorkBk.AddMainFlowSheet(FolderStructure.DirT0Tx, roomFlow);
            TestProgram.T0TxIgxlWorkBk.AddMainFlowSheet(FolderStructure.DirT0Tx, hotFlow);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            Assert.AreEqual(1, TestProgram.T0TxIgxlWorkBk.JobListSheets.Count);
            JobListSheet t0tx = TestProgram.T0TxIgxlWorkBk.JobListSheets.Values.First();
            Assert.AreNotEqual(null, t0tx.GetRow("FT1"));
            Assert.AreNotEqual(null, t0tx.GetRow("FT2"));
        }

        // ------------------------------------------------------------
        // 08. CustomPath branch executes with empty directory
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Run_CustomPath_Branch()
        {
            // Arrange
            string custom = Path.Combine(_root, "CustomFolder");
            Directory.CreateDirectory(custom);
            LocalSpecs.CustomPath = new List<string> { custom };

            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT1" } }
            };

            MainFlow flow = new MainFlow("Main_Flow_FT1") { JobNames = [] };
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            Assert.AreEqual(1, TestProgram.IgxlWorkBk.JobListSheets.Count);
        }

        // ------------------------------------------------------------
        // 09. TestProgramDefRow mapping and sheet name suffix
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Respect_TestProgramDefRow()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "FT", new List<string> { "FT1" } }
            };

            MainFlow flow = new MainFlow("Main_Flow_FT1") { JobNames = [] };
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);

            TestProgramRow def = new TestProgramRow
            {
                ProgramName = "MyProg",
                JobMapping = new Dictionary<string, string> { { "FT1", "ALPHA" } }
            };

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow(def);

            // Assert
            Assert.AreEqual(1, TestProgram.IgxlWorkBk.JobListSheets.Count);
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            Assert.AreEqual("JobList_MyProg", sheet.Name);
            Assert.AreSame(sheet, def.JobListSheet);

            JobRow? row = sheet.GetRow("FT1");
            Assert.AreNotEqual(null, row);
            Assert.AreEqual("Main_Flow_ALPHA", row!.FlowTable);
        }

        // ------------------------------------------------------------
        // 10. PortMap override by JobInfoSheet. Only PortSet is needed.
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Override_PortMap_By_JobInfoSheet()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>>
            {
                { "CP", new List<string> { "CP1" } }
            };

            MainFlow flow = new MainFlow("Main_Flow_CP1") { JobNames = [] };
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);

            // Default port map with non empty PortSets.
            PortMapSheet pmDefault = new PortMapSheet("PortMap_Default");
            pmDefault.Rows.Add(new PortSet("P0"));
            TestProgram.IgxlWorkBk.AddPortMapSheet(FolderStructure.DirPorts, pmDefault);

            // UF port map with non empty PortSets.
            PortMapSheet pmUF = new PortMapSheet("PortMap_UF");
            pmUF.Rows.Add(new PortSet("P1"));
            TestProgram.IgxlWorkBk.AddPortMapSheet(FolderStructure.DirPorts, pmUF);

            // Inject JobInfoSheet with TesterType = UF.
            JobInfoSheet info = new JobInfoSheet();
            info.Rows.Add(new JobInfoRow { JobName = "CP1", TesterType = "UF" });
            SetJobInfo(info);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("CP1");
            Assert.AreNotEqual(null, row);
            Assert.AreEqual("PortMap_UF", row!.PortMap);
        }

        // ------------------------------------------------------------
        // 11. PinMap present should set jobRow.PinMap
        // ------------------------------------------------------------

        [TestMethod]
        public void WorkFlow_Should_Set_Concurrent_When_Name_Match_And_HasConcurrent_True()
        {
            // Arrange
            LocalSpecs.JobMap = new Dictionary<string, List<string>> { { "CP", new List<string> { "CP2" } } };

            MainFlow flow = new MainFlow("Main_Flow_CP2")
            {
                JobNames = [],
                HasConcurrent = true
            };
            TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);

            DcSpecSheet dc = new DcSpecSheet("DC_Any", [], null!);
            TestProgram.IgxlWorkBk.AddDcSpecSheet(FolderStructure.DirDcSpec, dc);

            JobListMain target = new JobListMain();

            // Act
            target.WorkFlow();

            // Assert
            JobListSheet sheet = TestProgram.IgxlWorkBk.JobListSheets.Values.First();
            JobRow? row = sheet.GetRow("CP2");
            Assert.AreNotEqual(null, row);
            Assert.AreEqual("Main_Flow_CP2", row!.FlowTable);
            Assert.AreEqual("Concurrent Sequence", row.ConcurrentSequence);
        }

        // Helper to inject JobInfoSheet for PortMap override tests.
        private static void SetJobInfo(JobInfoSheet jobInfoSheet)
        {
            TestPlanStatic._jobInfoSheet = jobInfoSheet;
        }
    }
}
