using System.Collections.Generic;
using System.Drawing;
using System.IO;

using BinCutScriptLib.Static;

using CommonLib.ErrorReport;

using FileDiffLib;

using IgxlLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib;

namespace BinCutScriptLib.Test
{
    [TestClass]
    public class BinCutScriptLibTest
    {
        private static readonly string _inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        private static readonly string _outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        private static readonly string _expectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            // MockService.Mock();
        }

        [TestMethod]
        public void BinCutScriptLib0214Cp1NpiCs()
        {
            string subName = "TE06";
            string subType = "0214cp1npics";
            string outputPath = Path.Combine(_outputPath, subName, subType);
            string expectPath = Path.Combine(_expectPath, subName, subType);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var job = new Job(nameof(EnumJob.CP1));
            List<string> files = [Path.Combine(_inputPath, subName, subType, "cp1_online_bincut_hvcc_merge_pass12_Log.txt")];
            string testProgramFilePath = Path.Combine(_inputPath, subName, subType, "Hidra_20250213_2.igxl");

            var controller = new BinCutScriptMain(Print)
            {
                BinCutFilePath = "",
                BinCutPostFilePath = "",
                TestPlanFilePath = "",
                TestProgramFilePath = testProgramFilePath,
                IdsDistributionFilePath = "",
                OutPutFolder = outputPath,
                PowerBinning = "",
                ProjectName = "",
                DataLogFiles = files,
                IsVoltageByProgram = false,
                Job = job
            };

            ErrorReportManager.ClearErrors();
            BinCutConfig.Initialize();
            BinCutConfig.IsDebugPrint = false;

            controller.Run();


            bool fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void BinCutScriptLib0220Cp1MpCs()
        {
            string subName = "TE06";
            string subType = "0220cp1mpcs";
            string outputPath = Path.Combine(_outputPath, subName, subType);
            string expectPath = Path.Combine(_expectPath, subName, subType);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var job = new Job(nameof(EnumJob.CP1));
            List<string> files = [Path.Combine(_inputPath, subName, subType, "Loglevel_MP_full_2.txt")];
            string testProgramFilePath = Path.Combine(_inputPath, subName, subType, "Hidra_20250219_2.igxl");

            var controller = new BinCutScriptMain(Print)
            {
                BinCutFilePath = "",
                BinCutPostFilePath = "",
                TestPlanFilePath = "",
                TestProgramFilePath = testProgramFilePath,
                IdsDistributionFilePath = "",
                OutPutFolder = outputPath,
                PowerBinning = "",
                ProjectName = "",
                DataLogFiles = files,
                IsVoltageByProgram = false,
                Job = job
            };

            ErrorReportManager.ClearErrors();
            BinCutConfig.Initialize();
            BinCutConfig.IsDebugPrint = false;

            controller.Run();

            bool fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void BinCutScriptLib0220Ft1MpCs()
        {
            string subName = "TE06";
            string subType = "0220ft1mpcs";
            string outputPath = Path.Combine(_outputPath, subName, subType);
            string expectPath = Path.Combine(_expectPath, subName, subType);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var job = new Job(nameof(EnumJob.CP1));
            List<string> files = [Path.Combine(_inputPath, subName, subType, "FT1_offline_datalog_0219_no_Power_Binning_no_decision_MP.txt")];
            string testProgramFilePath = Path.Combine(_inputPath, subName, subType, "Hidra_20250219_2.igxl");

            var controller = new BinCutScriptMain(Print)
            {
                BinCutFilePath = "",
                BinCutPostFilePath = "",
                TestPlanFilePath = "",
                TestProgramFilePath = testProgramFilePath,
                IdsDistributionFilePath = "",
                OutPutFolder = outputPath,
                PowerBinning = "",
                ProjectName = "",
                DataLogFiles = files,
                IsVoltageByProgram = false,
                Job = job
            };

            ErrorReportManager.ClearErrors();
            BinCutConfig.Initialize();
            BinCutConfig.IsDebugPrint = false;

            controller.Run();

            bool fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void BinCutScriptLib0220Ft1NpiCs()
        {
            string subName = "TE06";
            string subType = "0220ft1npics";
            string outputPath = Path.Combine(_outputPath, subName, subType);
            string expectPath = Path.Combine(_expectPath, subName, subType);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            var job = new Job(nameof(EnumJob.CP1));
            List<string> files = [Path.Combine(_inputPath, subName, subType, "FT1_offline_datalog_0219_no_Power_Binning_no_decision_NPI.txt")];
            string testProgramFilePath = Path.Combine(_inputPath, subName, subType, "Hidra_20250219_2.igxl");

            var controller = new BinCutScriptMain(Print)
            {
                BinCutFilePath = "",
                BinCutPostFilePath = "",
                TestPlanFilePath = "",
                TestProgramFilePath = testProgramFilePath,
                IdsDistributionFilePath = "",
                OutPutFolder = outputPath,
                PowerBinning = "",
                ProjectName = "",
                DataLogFiles = files,
                IsVoltageByProgram = false,
                Job = job
            };

            ErrorReportManager.ClearErrors();
            BinCutConfig.Initialize();
            BinCutConfig.IsDebugPrint = false;

            controller.Run();

            bool fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        //[TestMethod]
        //public void BinCutScriptLib0227Cp1NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0227cp1npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "CP1_NoBankRead_COF_stored_nop_Decision_single_site.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250225_1.igxl");
        //    var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = testPlanFilePath;
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0227Cp2NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0227cp2npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "CP2_Online_Night.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250226_1.igxl");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0305Cp1NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0305cp1npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "Modify_C0_patt_manually_pass_20250303.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250226_2_V1.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0307Cp1NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0307cp1npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "WithoutBankRead_ErrorInHarvestEfuseWrite.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250306_1.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0307Cp2NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0307cp2npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "CP2_Online_Config_Flagstatus_modify_20250306.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250306_1.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0311Cp1NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0311cp1npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "CP1_Online_20250310.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250310_1.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0402Cp1DebugNpics()
        //{
        //    var subName = "TE06";
        //    var subType = "0402cp1debugnpics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "UR85_25C_N80CN7_W6_250402_V01A_CP1_PROD_PipeClean_MultiSite_Log_X6Y11.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "UR85_CP1_X4_E01_250401_V01A_CORR_BCV0P04_25C.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = true;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0409Cp1Npics()
        //{
        //    var subName = "TE06";
        //    var subType = "0409cp1npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "Hidra_250321_CP1_CORR_E804S9W01_X12Y20.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "TE06_CP1_X4_E06_250321_V06B_CORR_BCV1P83_CSHARP_KOM_25C.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}


        //[TestMethod]
        //public void BinCutScriptLib0409Cp1NpicsWithIds()
        //{
        //    var subName = "TE06";
        //    var subType = "0409cp1npics2";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "Hidra_250325_CP1_CORR_E804S9W01_X4Y12.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "TE06_CP1_X4_E06_250321_V06B_CORR_BCV1P83_CSHARP_KOM_25C.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}


        //[TestMethod]
        //public void BinCutScriptLib0307Ft1MpCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0307ft1mpcs";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "FT1_0306_doall.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250306_1.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "Hidra_C0_TestPlan_V06A_X_PF789_03011555.xlsx");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLib0219Ft1NpiCs()
        //{
        //    var subName = "TE06";
        //    var subType = "0219ft1npics";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "FT1_offline_datalog_0218.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "Hidra_20250218_1.igxl");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        //[TestMethod]
        //public void BinCutScriptLibVba()
        //{
        //    var subName = "TE06";
        //    var subType = "vba";
        //    var outputPath = Path.Combine(OutputPath, subName, subType);
        //    var expectPath = Path.Combine(ExpectPath, subName, subType);

        //    var job = new Job(EnumJob.CP1.ToString());
        //    List<string> files = new List<string>() { Path.Combine(InputPath, subName, subType, "TE06_CP1_X8_E04_V04A_PROD_VBA.txt") };
        //    var testProgramFilePath = Path.Combine(InputPath, subName, subType, "TE06_CP1_X8_E04_V04A_PROD_202412051400.igxl");
        //    //var testPlanFilePath = Path.Combine(InputPath, subName, subType, "");

        //    var controller = new BinCutScriptMain(print);
        //    controller.BinCutFilePath = "";
        //    controller.BinCutPostFilePath = "";
        //    controller.TestPlanFilePath = "";
        //    controller.TestProgramFilePath = testProgramFilePath;
        //    controller.IdsDistributionFilePath = "";
        //    controller.OutPutFolder = outputPath;
        //    controller.PowerBinning = "";
        //    controller.ProjectName = "";
        //    controller.DataLogFiles = files;
        //    controller.IsVoltageByProgram = false;
        //    controller.Job = job;

        //    ErrorReportManager.ResetErrors();
        //    BinCutConfig.Initialize();
        //    BinCutConfig.IsDebugPrint = false;

        //    controller.Run();

        //    var fail = new FileComparisonReport("BinCutScriptLib_" + subName + "_" + subType).IsFail(outputPath, expectPath, true);
        //    if (fail)
        //    {
        //        Assert.Fail("Unit Test Fail!!!");
        //    }
        //}

        private void Print(string text, Color color)
        {

        }
    }
}
