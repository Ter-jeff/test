using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.Extension;

using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using LogLib.Base;
using LogLib.Static;

using OfficeOpenXml;

using RfLib.Dvdc.GenTemplate.Bussiness;
using RfLib.Dvdc.GenTemplate.TestPlanFormat;
using RfLib.Dvdc.Reader.CapturePostProcess;
using RfLib.Dvdc.Reader.DsscSetup;

using TestPlanLib.Basic;
using TestPlanLib.Basic.Relay;
using TestPlanLib.Static;

namespace RfLib.Dvdc.GenTemplate
{
    public partial class TemplateGeneratorMain(HardIpInputData hardIpInputData)
    {
        [GeneratedRegex("^Channel", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("HardIP_|Wireless_|LCD_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        public string SheetName = "";
        private ChannelMapSheet? _channelMap;
        private readonly List<ChannelMapSheet> _multiChannelMapSheets = [];
        public Dictionary<string, List<PostProcessSheetRow>> Cpps = [];
        //public Dictionary<string, List<HardIpPattern>> TemplatePatternsDictionary = new Dictionary<string, List<HardIpPattern>>();
        public HardIpInputData HardIpInputData { get; } = hardIpInputData;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hardIpParaData"></param>
        /// <param name="isSplitDssc">if true, Split bits when DSSC Out larger than 24 bits</param>
        /// <param name="progresss"></param>
        public void WorkFlow(HardIpParaData hardIpParaData, bool isGenerateProgram, IProgress<Progress>? progresss = null)
        {
            string basicConfig = string.Format("{0}\\Settings\\Basic\\Basic_Configure_{1}.xlsx", LocalSpecs.SettingFolder, LocalSpecs.CurrentProject);
            var relayList = new RelayTableNew();

            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                basicConfig = string.Format("{0}\\Settings\\Basic_Configure_Default_RF.xlsx", LocalSpecs.SettingFolder);
            }
            else
            {
                basicConfig = string.Format("{0}\\Settings\\Basic_Configure_Default.xlsx", LocalSpecs.SettingFolder);
            }

            var inputExcel = new ExcelPackage(new FileInfo(basicConfig));
            SettingStatic.BasicConfigWorkbook = inputExcel.Workbook;
            Response.Report(string.Format("Loading...{0}", basicConfig), percentage: 1);

            //Kimi
            EpWorkbook.TestPlanWorkbook = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName)).Workbook;
            EpWorkbook.ScghWorkbook = new ExcelPackage(new FileInfo(LocalSpecs.ScghFileName)).Workbook;
            //

            NeededSheets.InitSheetName(LocalSpecs.Options.Device, LocalSpecs.SettingFolder, LocalSpecs.CurrentProject, Path.Combine(LocalSpecs.SettingFolder, "Settings", "NeedSheetConfig_Default.xml"));

            MultiTestSettingSheetsSingleton.Initialize(EpWorkbook.TestPlanWorkbook, LocalSpecs.CurrentProject, [.. LocalSpecs.JobMap.SelectMany(x => x.Value)], LocalSpecs.Options.Device);
            ErrorReportManager.ClearErrors();
            try
            {
                //Read SCGH file and patInfo
                //var readFlowMain = new ReadFlowMain();
                //readFlowMain.ReadInputFiles(para);
                var inputReader = new InputReader();
                var registerSheet = new Dictionary<EnumVbtFuncType, DsscSetupSheet>();
                DataTable? instrumentSetupSheet = new();
                var missingReg = new Dictionary<string, List<string>>();
                var dupRegName = new Dictionary<string, List<string>>();

                if (LocalSpecs.TestPlanFileName != "N/A")
                {
                    using var excel = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName));
                    foreach (ExcelWorksheet sheet in excel.Workbook.Worksheets)
                    {
                        if (MyRegex().IsMatch(sheet.Name))
                        {
                            ReadChanMapSheet readChanMapSheet = new ReadChanMapSheet();
                            _channelMap = readChanMapSheet.ReadSheet(sheet);
                            _multiChannelMapSheets.Add(_channelMap);
                        }
                    }
                }

                inputReader.ReadInput();

                Dictionary<string, PatternData> patternList = [];
                if (File.Exists(LocalSpecs.PatternListCsvFileName))
                {
                    Response.Report(string.Format("Reading PatternList file ..."), percentage: Convert.ToInt32(30));
                    patternList = PatternListReader.GetPatternListDic(LocalSpecs.PatternListCsvFileName) ?? [];
                }

                //Generate template row objects
                string outString = string.Format("Generating template sheet ...");
                Response.Report(outString, percentage: Convert.ToInt32(50));
                var templateAutoGen = new TemplateAutoGen(HardIpInputData, _multiChannelMapSheets, relayList)
                {
                    Cpps = Cpps
                };
                Dictionary<string, List<TemplateRow>> templates;
                if (LocalSpecs.Options.Device == EnumDevice.LCD)
                {
                    List<ScghLib.Reader.ProdCharSheetRow> scanItems = ScghStatic.ScghData.GetScanPatterns;
                    scanItems.ForEach(p => p.Block = "Scan");
                    List<ScghLib.Reader.ProdCharSheetRow> bistItems = ScghStatic.ScghData.GetBistPatterns;
                    bistItems.ForEach(p => p.Block = "MBist");
                    ScghStatic.ScghData.ConvertedPatternRowListByHardip.AddRange(scanItems);
                    ScghStatic.ScghData.ConvertedPatternRowListByHardip.AddRange(bistItems);
                }
                if (ScghStatic.ScghData.ConvertedPatternRowListByHardip.Count > 0)
                {
                    missingReg = TemplateAutoGenHelpers1.CheckMissingReg();
                    dupRegName = TemplateAutoGenHelpers1.CheckDupRegName();
                    registerSheet = new DsscRegisterTableBuilder(templateAutoGen).CreateRegisterTable();
                    instrumentSetupSheet = new DsscRegisterTableBuilder(templateAutoGen).CreateInstrumentSetup();
                    templates = new ScghTemplateRowGenerator(templateAutoGen).GenTemplatesByScgh(inputReader.PlanItems);
                }
                //else if (patternList.Count > 0)
                //{
                //    templates = templateAutoGen.GenTemplatesByPatList(SheetName, patternList);
                //}
                else
                {
                    outString = string.Format("There is no pattern , the template generation is terminated ...");
                    Response.Report(outString, EnumMessageLevel.Error, Convert.ToInt32(100));
                    return;
                }

                outString = string.Format("Writing TestPlan template file ...");
                Response.Report(outString, percentage: Convert.ToInt32(80));
                var templateWrite = new TemplateWriter(HardIpInputData)
                {
                    RegisterTable = registerSheet,
                    MissingReg = missingReg,
                    DupRegName = dupRegName,
                    InstrumentSetupTable = instrumentSetupSheet,
                    FuseInfos = templateAutoGen.ReferenceSetEfuseItems
                };
                templateWrite.WriteTemplate(templates);

                #region marked auto generate HardIP and RF block's Instance and Testflow sheet
                if (isGenerateProgram)
                {
                    var hardIpReader = new TestPlanConverter(ScghStatic.ScghData);
                    ExcelWorkbook planWBook = new ExcelPackage(new FileInfo(templateWrite.Filename)).Workbook;
                    var planSheet = new List<ExcelWorksheet>();
                    foreach (ExcelWorksheet ws in planWBook.Worksheets)
                    {
                        if (Regex.IsMatch(ws.Name, "DSSCSetupSheet|CaptoEfuseSetup|InstrumentSetup|CPP_|" + NeededSheets.HardIRegAssign, RegexOptions.IgnoreCase))
                        {
                            if (
                                TestProgram.NonIgxlSheetsList.SheetList.Exists(p => Path.GetFileNameWithoutExtension(p).EqualsIgnoreCase(ws.Name)))
                            {
                                continue;
                            }

                            ws.ExportWorkBook2Txt(LocalSpecs.TarFolder);
                            TestProgram.NonIgxlSheetsList.Add(LocalSpecs.TarFolder, ws.Name);
                        }
                        else if (MyRegex1().IsMatch(ws.Name))
                        {
                            planSheet.Add(ws);
                        }
                    }
                    HardIpInputData.PlanDic = new TestPlanConverter(ScghStatic.ScghData).ReadSheet(planSheet, null);
                    //TemplatePatternsDictionary = hardIpReader.ReadHardipPatterns(para.TestPlanSheetList, planWBook, ScghStatic.ScghData, null);
                }
                outString = string.Format("Generate TestPlan template successfully ...");

                Response.Report(outString, EnumMessageLevel.EndPoint, Convert.ToInt32(100));
                #endregion
            }
            catch (Exception e)
            {
                string outString = string.Format("Generate TestPlan template failed ..." + e.StackTrace);
                Response.Report(outString, EnumMessageLevel.Error, Convert.ToInt32(0));
            }
        }
    }
}
