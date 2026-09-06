using System;
using System.Collections.Generic;
using System.Linq;

using Automation;
using Automation.GenerateIgxl.Basic.Business;
using Automation.GenerateIgxl.Basic.Business.GenAc;
using Automation.GenerateIgxl.Basic.Business.GenNwire.Base;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.Singleton;
using Automation.Static;
using Automation.Static.Result;

using CommonLib.Enums;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;
using TestPlanLib.Static;

using ParaData = Automation.PreCheck.AllParaData.ParaData;

namespace LcdLib.Basic
{
    public class AutogenMainLcd : AutogenMainAp
    {
        private IoLevelsSheet? _ioLevels;

        #region Member Function
        public override void WorkFlow(ParaData paraData)
        {
            ProcessStatus = 10;
            try
            {
                BasicInitial();

                GenLevel();

                GenGlobalSpec();

                GenDcSpec();

                GenPatternSet();

                GenTimeSet();

                GenAcSpec();

                GenUfInstanceSheet();

                GenContinuity();

                GenNwire();

                GenBinTable();

                AddIgxlSheet();

                GenRelay();

                Response.Report("Basic Completed!", percentage: ProcessStatus = 100);
            }
            catch (Exception e)
            {
                string message = "Basic AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        protected override void BasicInitial()
        {
            Response.Report("Initializing Basic ...", percentage: ProcessStatus = 10);
            Initial = new BasicInitial();
            if (BasicResult.PatternList)
            {
                Response.Report("Reading Patterns/Timing Set from PatList File ...", percentage: ProcessStatus = 12);
                PatList = [.. (PatternListReader.GetPatternListDic(LocalSpecs.PatternFolder) ?? []).Select(x => x.Value)];
            }

            CopyNwireFile();
            try
            {
                MultiTestSettingSheetsSingleton = MultiTestSettingSheetsSingleton.Instance();
                ExcelWorksheet excelWorksheet = EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.IoLevels];
                if (excelWorksheet != null)
                {
                    _ioLevels = GetIoLevels(excelWorksheet);
                }
                //IoContiSheet = domainData.IoConti;
            }
            catch (Exception e)
            {
                Response.Report("Meet an Error in New Basic: " + e.StackTrace, EnumMessageLevel.Error, 100);
            }
        }

        private void GenDcSpec()
        {
            Response.Report("Generating DcSpec ...", percentage: ProcessStatus = 50);
            Mid.MultiDcSpecSheets = LcdDcSpecSheetsBuilder.Generate(MultiTestSettingSheetsSingleton, _ioLevels!, Mid.GlbSpecSheet);
        }

        private static IoLevelsSheet? GetIoLevels(ExcelWorksheet excelWorksheet)
        {
            IoLevelsSheetReader ioLevelsSheetReader = new IoLevelsSheetReader();
            return ioLevelsSheetReader.ReadSheet(excelWorksheet);
        }

        protected override void GenGlobalSpec()
        {
            Response.Report("Generating Global_SPEC ...", percentage: ProcessStatus = 35);
            Mid.GlbSpecSheet = LcdGlobalSpecBuilder.Generate(BasicInputData, MultiTestSettingSheetsSingleton, _ioLevels!);
        }

        protected override void GenAcSpec()
        {
            if (BasicResult.Ac)
            {
                //Initial AC Specs
                Response.Report("Generate AC Specs sheet ...", percentage: ProcessStatus = 45);
                Mid.AcInputSheet = Initial.InitalAcSymbols();
                var acGenerator = new AcGeneratorLcd(Mid.AcInputSheet);
                Mid.AcSpecSheet = acGenerator.GenerateFlow(Mid.TimeSetSheets, PatList);
            }
        }

        protected override void GenTimeSet()
        {
            if (BasicResult.TimeSet)
            {
                Response.Report($"Copying Timing Set from path {LocalSpecs.TimeSetFolder} ...", percentage: ProcessStatus = 72);
                try
                {
                    var tsetGenerator = new TimeSetGenerator();
                    Mid.TimeSetSheets = tsetGenerator.GenerateFlow(PatList, LocalSpecs.TimeSetFolder, FolderStructure.DirTimings);

                    if (Mid.TimeSetSheets.Count == 0)
                    {
                        Response.Report("TimeSet Path not exised, or found no TimeSet File!", EnumMessageLevel.Warning, ProcessStatus = 75);
                    }

                    if (tsetGenerator.TimeSetWithWrongForamtRows.Count != 0)
                    {
                        foreach (KeyValuePair<string, List<int>> timeSet in tsetGenerator.TimeSetWithWrongForamtRows)
                        {
                            Response.Report($"TSet Row Wrong Format: {timeSet.Key}, @RowNum: {string.Join(",", timeSet.Value)}, Compensate empty data", EnumMessageLevel.Error, ProcessStatus = 75);
                        }
                    }

                    //Update MCG Mode
                    TimeSetMcgMode mcgMode = new TimeSetMcgMode(NwireSingleton.Instance().SettingInfo.NwirePins);
                    mcgMode.ConverFlow(Mid.TimeSetSheets, PatList);

                    TimeSetPlus timeSetPlus = new TimeSetPlus(NwireSingleton.Instance().NonFrcSetting);
                    timeSetPlus.PlusFlow(Mid.TimeSetSheets);

                }
                catch (Exception ex)
                {
                    Response.Report("Generating Timing Set failed! " + ex.Message, EnumMessageLevel.Warning, ProcessStatus = 90);
                }
            }
            else
            {
                Mid.TimeSetSheets = [];
                TimeSetPlus timeSetPlus = new TimeSetPlus(NwireSingleton.Instance().NonFrcSetting);
                timeSetPlus.PlusFlow(Mid.TimeSetSheets);
            }
        }

        protected override void GenNwire()
        {
            NwireResult? nwireResult = LcdNwireResultBuilder.Generate();
            if (nwireResult != null)
            {
                NwireResult = nwireResult;
            }
        }

        private static void GenBinTable()
        {
            LcdBinTableSystemErrorRow.Add();
        }
        #endregion
    }
}
