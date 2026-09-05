using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using TestPlanLib;
using TestPlanLib.Static;

namespace BinCutScriptLib
{
    public class BinCutScriptMain(Action<string, Color> richTextBoxAppend)
    {
        public string BinCutFilePath = "";
        public string BinCutPostFilePath = "";
        public string ModeSequenceFilePath = "";
        public string TestPlanFilePath = "";
        public string TestProgramFilePath = "";
        public string IdsDistributionFilePath = "";
        public string PatDashFilePath = "";
        public string PowerBinning = "";
        public string ScghFilePath = "";
        public string DataLogFolder = "";
        public string OutPutFolder = "";
        public string ProjectName = "";
        public string TempFolder { get { return Path.Combine(OutPutFolder, "Temp"); } }
        public bool? IsVoltageByProgram { get; set; }
        public List<string> DataLogFiles { get; set; } = [];
        public Job Job { get; set; } = BinCutData.Job;
        public bool CmdMode;

        public Action<string, Color> RichTextBoxAppend = richTextBoxAppend;

        public void Run()
        {
            BinCutConfig.GetProjectConfig(ProjectName);
            NeededSheets.InitSheetName(EnumDevice.AP, Directory.GetCurrentDirectory(), ProjectName);
            if (new BinCutReadMain(this, RichTextBoxAppend).ReadBinCutData(true))
            {
                DatalogValidation();
            }

            Directory.Delete(TempFolder, true);
        }

        public void DatalogValidation()
        {
            #region If input file be locked
            string tempBinCutFile = "";
            string tempTestPlanFile = "";
            string tempIdsDistributionFile = "";
            string tempScghFile = "";
            if (StringExtensions.IsOpened(BinCutFilePath))
            {
                tempBinCutFile = BinCutScriptMainHelpers.CopyToTemp(BinCutFilePath);
                BinCutFilePath = tempBinCutFile;
            }
            if (StringExtensions.IsOpened(BinCutPostFilePath))
            {
                string tempBinCutPostFile = BinCutScriptMainHelpers.CopyToTemp(BinCutPostFilePath);
                BinCutPostFilePath = tempBinCutPostFile;
            }
            if (StringExtensions.IsOpened(TestPlanFilePath))
            {
                tempTestPlanFile = BinCutScriptMainHelpers.CopyToTemp(TestPlanFilePath);
                TestPlanFilePath = tempTestPlanFile;
            }
            if (StringExtensions.IsOpened(IdsDistributionFilePath))
            {
                tempIdsDistributionFile = BinCutScriptMainHelpers.CopyToTemp(IdsDistributionFilePath);
                IdsDistributionFilePath = tempIdsDistributionFile;
            }
            if (StringExtensions.IsOpened(ScghFilePath))
            {
                tempScghFile = BinCutScriptMainHelpers.CopyToTemp(ScghFilePath);
                ScghFilePath = tempScghFile;
            }
            if (StringExtensions.IsOpened(PowerBinning))
            {
                tempScghFile = BinCutScriptMainHelpers.CopyToTemp(PowerBinning);
                PowerBinning = tempScghFile;
            }
            #endregion

            bool csFlag = AlgorithmBaseHelpers.CheckCsharpLogFormat(DataLogFiles.First());
            if (!csFlag)
            {
                var checker = new BinCutCheckMain(this, BinCutData.PinInfos, BinCutData.PowerPins, BinCutData.BinningTables, BinCutData.BinCutFlowTables);
                checker.Execute();
            }
            else
            {
                var checker = new BinCutCheckMainCs(this, BinCutData.PinInfos, BinCutData.PowerPins, BinCutData.BinningTables, BinCutData.BinCutFlowTables);
                checker.Execute();
            }

            BinCutScriptMainHelpers.RemoveTemp(tempBinCutFile);
            BinCutScriptMainHelpers.RemoveTemp(tempTestPlanFile);
            BinCutScriptMainHelpers.RemoveTemp(tempIdsDistributionFile);
            BinCutScriptMainHelpers.RemoveTemp(tempScghFile);
        }
    }
}
