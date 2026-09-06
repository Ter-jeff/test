using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.SetFunction;
using BinCutScriptLib.Static;

using CommonLib.ErrorReport;
using CommonLib.Extension;

using LogLib.Utility;

using TestPlanLib;

using PinInfo = TestPlanLib.BinCut.Binning.PinInfo;

namespace BinCutScriptLib
{
    public abstract class AlgorithmBase
    {
        public static EnumBcConfig Method
        {
            get
            {
                if (BinCutConfig.VddbinCofStepInheritanceNewLogic.Equals(true))
                {
                    return EnumBcConfig.Vddbin_COF_StepInheritance_New_Logic;
                }

                if (BinCutConfig.VddbinCofStepInheritance.Equals(true))
                {
                    return EnumBcConfig.Vddbin_COF_StepInheritance;
                }

                if (BinCutConfig.DebugBinCutCofStored.Equals(true))
                {
                    return EnumBcConfig.Debug_BinCutCOF_Stored;
                }

                if (!string.IsNullOrEmpty(BinCutConfig.GradeSearchMethodSelected))
                {
                    if (BinCutConfig.GradeSearchMethodSelected == "Conventional")
                    {
                        return EnumBcConfig.Conventional;
                    }

                    if (BinCutConfig.GradeSearchMethodSelected == "FollowingModes")
                    {
                        return EnumBcConfig.FollowingModes;
                    }

                    if (BinCutConfig.GradeSearchMethodSelected == "PerformanceModes")
                    {
                        return EnumBcConfig.PerformanceModes;
                    }

                    if (BinCutConfig.GradeSearchMethodSelected == "BinningDomains")
                    {
                        return EnumBcConfig.BinningDomains;
                    }
                }
                BinCutConfig.GradeSearchMethodSelected = nameof(EnumBcConfig.Conventional);
                return EnumBcConfig.Conventional;
            }
        }

        public int ErrorCount;

        public List<SiteInfo> CurrentDiceInfos = [];
        public List<List<SiteInfo>> TotalDiceInfos = [];

        public string CurInstanceName = "";
        protected string OutputFile = string.Empty;
        public Job Job = BinCutData.Job;
        public CheckManager CheckManager;
        public CheckManager CheckManagerTotal;
        public OneTouchDown OneTouchDown = new();
        public int LineNoCounter;

        public int ActiveSiteCount;
        public int TchCnt;

        public List<Alarm> AlarmList = [];
        public List<BvName> BvNames = [];
        public List<BvName> BvNamesBackup = [];

        protected AlgorithmBase(BinCutScriptMain binCutScriptMain)
        {
            BinCutController.Controller = binCutScriptMain;
            CheckManager = new CheckManager();
            CheckManagerTotal = new CheckManager();
        }

        public virtual void Execute()
        {
            CheckManagerTotal.Initialize();
            foreach (string file in BinCutController.Controller.DataLogFiles)
            {
                CheckManager.Initialize();
                LineNoCounter = 0;
                string fileName = Path.GetFileNameWithoutExtension(file);
                OutputFile = Path.Combine(BinCutController.Controller.OutPutFolder, "Check_result_" + fileName + ".txt");
                CurrentDiceInfos = [];
                try
                {
                    if (!Workflow(file))
                    {
                        return;
                    }

                    TotalDiceInfos.Add(CurrentDiceInfos);
                    CheckManagerTotal.Add(CheckManager);
                    PrintReportMain.PrintReport(BinCutController.Controller.RichTextBoxAppend, [CurrentDiceInfos], OutputFile, AlarmList, Job, CheckManager, BinCutController.Controller.TempFolder);
                }
                catch (Exception e)
                {
                    if (BinCutController.Controller.CmdMode)
                    {
                        if (CurrentDiceInfos.Count == 0)
                        {
                            BinCutController.Controller.RichTextBoxAppend("The error message is: Cannot found bincut data in log " + file, Color.Red);
                        }
                        else
                        {
                            BinCutController.Controller.RichTextBoxAppend(e.Message + Environment.NewLine + e.StackTrace, Color.Red);
                        }

                        throw;
                    }
                    else
                    {
                        ErrorMessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "");
                    }
                }
            }

            if (BinCutController.Controller.DataLogFiles.Count != 1)
            {
                OutputFile = Path.Combine(BinCutController.Controller.OutPutFolder, "Check_result_Total");
                CheckManager = CheckManagerTotal;
                BinCutController.Controller.RichTextBoxAppend("*******************************************************************************************", Color.Blue);
                PrintReportMain.PrintReport(BinCutController.Controller.RichTextBoxAppend, TotalDiceInfos, OutputFile, AlarmList, Job, CheckManager, BinCutController.Controller.TempFolder);
            }
            ErrorReportManager.ClearErrors();
        }

        protected void PrintCurrentSite(SiteInfo[] siteInfoArray)
        {
            BinCutController.Controller.RichTextBoxAppend("Touch down site = " + ActiveSiteCount, Color.Blue);
            for (int i = 0; i < siteInfoArray.Length; i++)
            {
                if (siteInfoArray[i].Site != -1)
                {
                    string msg =
                        $"  Current active site = {siteInfoArray[i].Site} (XCoor={siteInfoArray[i].XCoor:D3}, YCoor={siteInfoArray[i].YCoor:D3}, Bin={siteInfoArray[i].SortBin})";
                    BinCutController.Controller.RichTextBoxAppend(msg, Color.Blue);
                }
            }
        }

        protected virtual bool GetOneTouchDown(StreamReader streamReader, ref OneTouchDown oneTouchDown, ref SiteInfo[] siteInfoArray)
        {
            bool isOneTouchDown = false;
            bool isFoundSortBin = false;
            bool beforeBVfalg = true;
            bool beforeoutsideBVfalg = true;
            bool debugLines = false;

            oneTouchDown.Lines.Clear();
            //STEP1. Get test flow start pos, and calc activeSites
            string? line = "";
            while (!streamReader.EndOfStream)
            {
                line = streamReader.ReadLine();
                LineNoCounter++;
                if (line != null && line.Contains("Device#"))
                {
                    isOneTouchDown = true;
                    break;
                }
            }
            if (!isOneTouchDown)
            {
                return false;
            }

            ActiveSiteCount = AlgorithmBaseHelpers.GetDevice(oneTouchDown, line!, LineNoCounter);

            AlgorithmBaseHelpers.ReadUntilEndOfOneTouch(streamReader, oneTouchDown, ref isFoundSortBin, ref beforeBVfalg, ref beforeoutsideBVfalg, ref debugLines, ref LineNoCounter, ref AlarmList);

            if (!isFoundSortBin)
            {
                return false;
            }

            AlgorithmBaseHelpers.GetSortBin(streamReader, oneTouchDown, siteInfoArray, ref LineNoCounter);

            string? line1 = streamReader.ReadLine();
            LineNoCounter++;
            oneTouchDown.Lines.Add(new BinCutLineBase { Line = line1!, LineNo = LineNoCounter });
            AlgorithmBaseHelpers.GetCoor(streamReader, oneTouchDown, siteInfoArray, ref LineNoCounter);

            return true;
        }

        protected SiteInfo[] InitSite(int maxSite, List<PinInfo> pinInfos)
        {
            var allDice = new SiteInfo[maxSite];
            for (int site = 0; site < maxSite; site++)
            {
                allDice[site] = new SiteInfo { Site = -1, Sort = -1, SortBin = -1, XCoor = -1, YCoor = -1, Bin = 1, HarvBin = [], TuchNum = TchCnt };
                allDice[site].InitDiceByPowers(pinInfos);
                allDice[site].InitHarvFlag();
            }

            return allDice;
        }

        protected bool Workflow(string dataLogFile)
        {
            TchCnt = 0;
            try
            {
                BinCutController.Controller.RichTextBoxAppend($"File Name: {Path.GetFileName(dataLogFile)}", Color.Blue);
                bool csFlag = AlgorithmBaseHelpers.CheckCsharpLogFormat(dataLogFile);
                if (!DataLogReader.ReadJob(dataLogFile, out int maxSite, out string programName, out Job, BinCutController.Controller.RichTextBoxAppend))
                {
                    using var sw = new StreamWriter(OutputFile);
                    BinCutPrint.PrintNoJobMessage(sw, "No Find Job Information in log.");
                    return false;
                }

                #region Collect BinCutConfig Information
                BinCutDatalogConfigReader.ReadBinCutDatalogConfig(dataLogFile, BinCutController.Controller.TempFolder, ref Job, programName, csFlag);
                #endregion

                Run(dataLogFile, maxSite);
                //<- 若欲疊合RT log, call此function

                int duplicateCnt = AlgorithmBaseHelpers.SuperposRtDice(ref CurrentDiceInfos);
                BinCutController.Controller.RichTextBoxAppend($"Totally {duplicateCnt} dice are duplicate.", duplicateCnt == 0 ? Color.Blue : Color.Orange);
            }
            catch (Exception e)
            {
                if (BinCutController.Controller.CmdMode)
                {
                    BinCutController.Controller.RichTextBoxAppend(e.Message + Environment.NewLine + e.StackTrace, Color.Red);
                }
                else
                {
                    if (e.Message.StartsWithIgnoreCase("Autogen"))
                    {
                        ErrorMessageBox.Show(e.Message, "");
                    }
                    else
                    {
                        ErrorMessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "");
                    }

                    return false;
                }
            }
            return true;
        }

        protected abstract void Run(string dataLogFile, int maxSite);

        protected void ProAction(SiteInfo[] siteInfoArray, List<PinInfo> pinInfos, StreamWriter streamWriter)
        {
            //Must have => Product_Identifier & product
            ReadDvfmManager.ReadProduct(streamWriter, Job.JobType, ref OneTouchDown, ref siteInfoArray);
            //Just keep the same data structure as CP1

            CreatePowerStepsMain.CreatePowerSteps(ref siteInfoArray, pinInfos, Job.JobType);

            AlgorithmBaseHelpers.ModifyStepByEfuseValue(siteInfoArray, BinCutController.Controller.CmdMode, BinCutController.Controller.RichTextBoxAppend);

            //Set non-CorePowerStep to 0 for powerbinning
            AlgorithmBaseHelpers.InitNonCorePwrStep(siteInfoArray, pinInfos);

            //Init sram core power step
            AlgorithmBaseHelpers.InitSramPwrStep(siteInfoArray, pinInfos);
        }
    }
}
