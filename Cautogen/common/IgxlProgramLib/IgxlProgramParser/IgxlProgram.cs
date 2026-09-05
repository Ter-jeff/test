using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Automation.Static;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using LogLib.Static;

using AcSpecSheet = IgxlLib.IgxlSheets.AcSpecSheet;
using BinTableSheet = IgxlLib.IgxlSheets.BinTableSheet;
using ChannelMapSheet = IgxlLib.IgxlSheets.ChannelMapSheet;
using DcSpecSheet = IgxlLib.IgxlSheets.DcSpecSheet;
using FlowSheet = IgxlLib.IgxlSheets.SubFlowSheet;
using GlobalSpec = IgxlLib.IgxlBase.GlobalSpec;
using GlobalSpecSheet = IgxlLib.IgxlSheets.GlobalSpecSheet;
using Pin = IgxlLib.IgxlBase.Pin;
using PinGroup = IgxlLib.IgxlBase.PinGroup;
using PinMapSheet = IgxlLib.IgxlSheets.PinMapSheet;
using PortMapSheet = IgxlLib.IgxlSheets.PortMapSheet;
using PortRow = IgxlLib.IgxlBase.PortRow;

namespace Cautogen.common.IgxlProgramLib.IgxlProgramParser
{
    public class IgxlProgram
    {
        private readonly string _selectJob;

        public List<BinTableSheet> BintableSheets = new List<BinTableSheet>();
        public List<AcSpecSheet> AcSpecSheets = new List<AcSpecSheet>();
        public List<FlowSheet> FlowSheetsAll = new List<FlowSheet>();
        public List<FlowSheet> FlowSheets = new List<FlowSheet>();
        public List<MainFlow> MainFlowSheets = new List<MainFlow>();
        public List<TimeSetBasicSheet> TimeSetSheets = new List<TimeSetBasicSheet>();
        public List<LevelSheet> LevelSheets = new List<LevelSheet>();
        public List<CharSheet> CharSheets = new List<CharSheet>();
        public List<PortMapSheet> PortMapSheets = new List<PortMapSheet>();
        public List<PatSetSheet> PatSetSheets = new List<PatSetSheet>();
        public List<PatSetSubSheet> PatSetSubSheets = new List<PatSetSubSheet>();
        public List<InstanceSheet> InstanceSheets = new List<InstanceSheet>();
        public JobListSheet JoblistSheet;
        public List<PinMapSheet> PinMaps;
        public List<Pin> PinList;
        public List<PinGroup> PinGroups;
        public List<ChannelMapSheet> ChannelMapSheets = new List<ChannelMapSheet>();
        public List<DcSpecSheet> DcSpecSheets = new List<DcSpecSheet>();
        public List<DcSpec> DcSpecDatas = new List<DcSpec>();
        public List<string> DcCategoryList = new List<string>();
        public GlobalSpecSheet GlbSpecSheet;
        public List<GlobalSpec> GlobalSpecRows = new List<GlobalSpec>();
        public List<PortRow> PortRows = new List<PortRow>();
        public List<string> FrcList = new List<string>();
        public List<string> AcSpecsSymbols = new List<string>();
        public List<ReferenceSheet> ReferenceSheets = new List<ReferenceSheet>();

        public IgxlProgram(string job)
        {
            _selectJob = job;
        }

        private Dictionary<string, List<FlowRow>> _flowNameRowDict;
        public Dictionary<string, List<FlowRow>> FlowNameRowDict
        {
            get
            {
                return _flowNameRowDict ?? (_flowNameRowDict = FlowSheets
                    .SelectMany(p => p.Rows)
                    .Where(a => !string.IsNullOrEmpty(a.Parameter))
                    .GroupBy(p => p.Parameter)
                    .ToDictionary(p => p.Key, p => p.ToList()));
            }
        }

        public void PrintOthers(string outputFolder, string exportProg, string otherFolder)
        {
            try
            {
                IEnumerable<string> files = Directory.GetFiles(exportProg, "*.*", SearchOption.AllDirectories)
                   .Where(x => x.EndsWith("txt", StringComparison.OrdinalIgnoreCase) ||
                               x.EndsWith("bas", StringComparison.OrdinalIgnoreCase) ||
                               x.EndsWith("cls", StringComparison.OrdinalIgnoreCase));

                var outfiles = Directory.GetFiles(outputFolder, "*.*", SearchOption.AllDirectories)
                    .Where(x => x.EndsWith("txt", StringComparison.OrdinalIgnoreCase) ||
                               x.EndsWith("bas", StringComparison.OrdinalIgnoreCase) ||
                               x.EndsWith("cls", StringComparison.OrdinalIgnoreCase))
                    .ToList().Select(x => Path.GetFileName(x).Replace("%20", " ")).ToList();
                foreach (string file in files)
                {
                    string fileName = Path.GetFileName(file).Replace("%20", " ");
                    if (outfiles.Exists(x => Path.GetFileName(x).Equals(Path.GetFileName(fileName), StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    if (fileName.Equals("IGLinkManifest.txt", StringComparison.CurrentCulture))
                    {
                        continue;
                    }

                    //if (fileName.Equals("Pattern_Subroutine.txt", StringComparison.CurrentCulture))
                    //    continue;

                    if (File.Exists(file))
                    {
                        if (!Directory.Exists(otherFolder))
                        {
                            Directory.CreateDirectory(otherFolder);
                        }

                        File.Copy(file, Path.Combine(otherFolder, Path.GetFileName(file)));
                    }
                }

            }
            catch
            {
                Response.Report($"Call From T-Autogen no needed to export sheet");
                return;
            }
        }

        public void Print(string outputFolder, string exportProg)
        {
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            if (PinMaps != null)
            {
                foreach (PinMapSheet pinMap in PinMaps)
                {
                    string name = pinMap.Name + ".txt";
                    string txt = Path.Combine(exportProg, name);
                    if (File.Exists(txt))
                    {
                        File.Copy(txt, Path.Combine(outputFolder, name));
                    }
                }
            }

            if (ChannelMapSheets != null)
            {
                foreach (ChannelMapSheet channelMapSheet in ChannelMapSheets)
                {
                    string name = channelMapSheet.Name + ".txt";
                    string txt = Path.Combine(exportProg, name);
                    if (File.Exists(txt))
                    {
                        File.Copy(txt, Path.Combine(outputFolder, name));
                    }
                }
            }

            foreach (TimeSetBasicSheet timeSetSheet in TimeSetSheets)
            {
                string name = timeSetSheet.Name + ".txt";
                timeSetSheet.Write(Path.Combine(outputFolder, name));
            }

            foreach (InstanceSheet instanceSheets in InstanceSheets)
            {
                string name = instanceSheets.Name + ".txt";
                instanceSheets.Write(Path.Combine(outputFolder, name));
            }

            foreach (DcSpecSheet dcSpecSheet in DcSpecSheets)
            {
                string name = dcSpecSheet.Name + ".txt";
                dcSpecSheet.Write(Path.Combine(outputFolder, name));
            }
        }

        public void LoadIgxlInstance(string path)
        {
            var readInstance = new Task(() =>
            {
                InstanceSheets = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTTestInstancesSheet)
                        .Cast<InstanceSheet>()
                        .ToList();
            });

            readInstance.Start();
            Task.WaitAll(readInstance);
        }

        public void LoadIgxlForInputdef(string path)
        {
            var readInstance = new Task(() =>
            {
                InstanceSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTTestInstancesSheet)
                        .Cast<InstanceSheet>()
                        .ToList();
            });

            var readPatSet = new Task(() =>
            {
                PatSetSheets =
                   IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPatternSetSheet)
                        .Cast<PatSetSheet>()
                        .ToList();
            });

            readInstance.Start();
            readPatSet.Start();
            Task.WaitAll(readInstance, readPatSet);
        }

        public void LoadIgxlForPreProcess()
        {
            InstanceSheets = TestProgram.IgxlWorkBk.InsSheets.Values.ToList();
            FlowSheets = TestProgram.IgxlWorkBk.SubFlowSheets.Values.ToList();
            FlowSheets.AddRange(TestProgram.IgxlWorkBk.MainFlowSheets.Values.ToList());
            FlowSheetsAll = FlowSheets;
            BintableSheets = TestProgram.IgxlWorkBk.BinTblSheets.Values.ToList();
            TimeSetSheets = TestProgram.IgxlWorkBk.TimeSetSheets.Values.ToList();
            PatSetSheets = TestProgram.IgxlWorkBk.PatSetSheets.Values.ToList();
            PatSetSubSheets = TestProgram.IgxlWorkBk.PatSetSubSheets.Values.ToList();
            DcSpecSheets = TestProgram.IgxlWorkBk.DcSpecSheets.Values.ToList();
            DcSpecDatas = DcSpecSheets.SelectMany(p => p.Rows).ToList();
            if (DcSpecDatas.Any())
            {
                DcCategoryList = DcSpecDatas[0].CategoryList.Select(p => p.Name).ToList();
            }

            AcSpecSheets = TestProgram.IgxlWorkBk.AcSpecSheets.Values.ToList();

            var joblist = TestProgram.IgxlWorkBk.JobListSheets.Values.ToList();
            if (joblist.Any())
            {
                JoblistSheet = joblist.First();
                JobRow userJob = JoblistSheet.Rows.FirstOrDefault(x => x.JobName.Equals(_selectJob, StringComparison.OrdinalIgnoreCase));
                string jobPinMap = "";
                if (userJob != null)
                {
                    jobPinMap = userJob.PinMap;
                }

                PinList = (List<Pin>)TestProgram.IgxlWorkBk.PinMapPair.Value.PinList;
                PinGroups = (List<PinGroup>)TestProgram.IgxlWorkBk.PinMapPair.Value.GroupList;
            }

            ChannelMapSheets = TestProgram.IgxlWorkBk.ChannelMapSheets.Values.ToList();
            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value != null)
            {
                GlobalSpecRows = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.Rows;
            }
            PortRows = TestProgram.IgxlWorkBk.PortMapSheets.Values.Cast<PortMapSheet>().SelectMany(p => p.Rows).SelectMany(q => q.PortRows).ToList();

            InstanceSheet testInstCommonSheet = InstanceSheets.FirstOrDefault(x => x.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            if (testInstCommonSheet != null)
            {
                IEnumerable<InstanceRow> frcInstRows = testInstCommonSheet.Rows.Where(x => x.TestName.StartsWith("FreeRunClk_Enable_", StringComparison.OrdinalIgnoreCase)
                                    && x.IsBackup == false);
                FrcList = frcInstRows.Select(x => Regex.Replace(x.TestName, @"FreeRunClk_Enable_", "", RegexOptions.IgnoreCase)).ToList();
            }

            AcSpecSheet acSpecsSheet = AcSpecSheets.FirstOrDefault(x => x.Name.Equals("AC_Specs", StringComparison.OrdinalIgnoreCase));
            if (acSpecsSheet != null)
            {
                var acSpecsSymbols = acSpecsSheet.Rows.Select(x => x.Symbol).ToList();
                acSpecsSymbols.RemoveAll(string.IsNullOrEmpty);
                acSpecsSymbols.RemoveAll(x => x.Replace(" ", "").Equals("UsingTSet"));
                AcSpecsSymbols = acSpecsSymbols;
            }
        }
        public void LoadIgxlForPreProcess(string path)
        {
            var readInstance = new Task(() =>
            {
                InstanceSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTTestInstancesSheet)
                        .Cast<InstanceSheet>()
                        .ToList();
            });

            var readFlow = new Task(() =>
            {
                FlowSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTFlowtableSheet)
                        .Cast<FlowSheet>()
                        .ToList();
                FlowSheetsAll = FlowSheets;
            });


            var readBinTable = new Task(() =>
            {
                BintableSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTBintablesSheet)
                    .Cast<BinTableSheet>()
                    .ToList();
            });

            var readLevel = new Task(() =>
            {
                LevelSheets =
                IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTLevelSheet)
                .Cast<LevelSheet>()
                .ToList();
            });

            #region read Time/PatternSet
            var readTimeset = new Task(() =>
            {
                TimeSetSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTTimesetBasicSheet)
                        .Cast<TimeSetBasicSheet>()
                        .ToList();
            });

            var readPattern = new Task(() =>
            {
                PatSetSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPatternSetSheet)
                        .Cast<PatSetSheet>()
                        .ToList();
            });


            var readPatSubroutine = new Task(() =>
            {
                PatSetSubSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPatternSubroutineSheet)
                        .Cast<PatSetSubSheet>()
                        .ToList();
            });
            #endregion

            #region read DC/AC
            var readDcCategory = new Task(() =>
            {
                DcSpecSheets = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTDCSpecSheet)
                    .Cast<DcSpecSheet>().ToList();
                DcSpecDatas = DcSpecSheets.SelectMany(p => p.Rows).ToList();
                if (DcSpecDatas.Any())
                {
                    DcCategoryList = DcSpecDatas[0].CategoryList.Select(p => p.Name).ToList();
                }
            });

            var readAcCategory = new Task(() =>
            {
                AcSpecSheets = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTACSpecSheet)
                        .Cast<AcSpecSheet>().ToList();
            });

            var readOthers = new Task(() =>
            {
                var joblist = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTJobListSheet)
                    .Cast<JobListSheet>().ToList();
                if (joblist.Any())
                {
                    JoblistSheet = joblist.First();
                }

                {
                    JobRow userJob = JoblistSheet.Rows.FirstOrDefault(x => x.JobName.Equals(_selectJob, StringComparison.OrdinalIgnoreCase));
                    string jobPinMap = "";
                    if (userJob != null)
                    {
                        jobPinMap = userJob.PinMap;
                    }

                    PinMaps = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPinMap)
                        .Cast<PinMapSheet>().ToList();
                    PinMapSheet pinMap = null;
                    if (!string.IsNullOrEmpty(jobPinMap))
                    {
                        pinMap = PinMaps.FirstOrDefault(x => x.Name.Equals(jobPinMap, StringComparison.OrdinalIgnoreCase));
                    }

                    if (pinMap == null)
                    {
                        pinMap = PinMaps[0];
                    }

                    PinList = (List<Pin>)pinMap.PinList;
                    PinGroups = (List<PinGroup>)pinMap.GroupList;
                }

                ChannelMapSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTChanMap)
                        .Cast<ChannelMapSheet>()
                        .ToList();

                GlbSpecSheet = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTGlobalSpecSheet).Cast<GlobalSpecSheet>().FirstOrDefault();
                GlobalSpecRows = GlbSpecSheet.Rows.ToList();

                PortMapSheets = IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPortMapSheet)
                                    .Cast<PortMapSheet>().ToList();
                PortRows = PortMapSheets.SelectMany(p => p.Rows).SelectMany(q => q.PortRows).ToList();

                ReferenceSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTReferencesSheet)
                        .Cast<ReferenceSheet>().ToList();
            });
            #endregion

            var readChar = new Task(() =>
            {
                CharSheets =
                IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTCharacterizationSheet)
                .Cast<CharSheet>()
                .ToList();
            });

            #region Task Start
            readDcCategory.Start();
            readAcCategory.Start();
            readBinTable.Start();
            readInstance.Start();
            readFlow.Start();
            readTimeset.Start();
            readLevel.Start();
            readChar.Start();
            readPattern.Start();
            readPatSubroutine.Start();
            readOthers.Start();
            #endregion

            Task.WaitAll(readBinTable, readInstance, readFlow, readTimeset, readLevel, readChar, readPattern, readPatSubroutine,
                readOthers, readDcCategory, readAcCategory);

            InstanceSheet testInstCommonSheet = InstanceSheets.FirstOrDefault(x => x.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            if (testInstCommonSheet != null)
            {
                IEnumerable<InstanceRow> frcInstRows = testInstCommonSheet.Rows.Where(x => x.TestName.StartsWith("FreeRunClk_Enable_", StringComparison.OrdinalIgnoreCase)
                                    && !x.IsBackup);
                FrcList = frcInstRows.Select(x => Regex.Replace(x.TestName, "FreeRunClk_Enable_", "", RegexOptions.IgnoreCase)).ToList();
            }

            AcSpecSheet acSpecsSheet = AcSpecSheets.FirstOrDefault(x => x.Name.Equals("AC_Specs", StringComparison.OrdinalIgnoreCase));
            if (acSpecsSheet != null)
            {
                var acSpecsSymbols = acSpecsSheet.Rows.Select(x => x.Symbol).ToList();
                acSpecsSymbols.RemoveAll(string.IsNullOrEmpty);
                acSpecsSymbols.RemoveAll(x => x.Replace(" ", "").Equals("UsingTSet"));
                AcSpecsSymbols = acSpecsSymbols;
            }

            List<FlowSheet> mFlowSheets = FlowSheets.Where(p => p.Name.Contains("Main_Flow")).ToList();
            foreach (FlowSheet sheet in mFlowSheets)
            {
                MainFlow flow = new MainFlow(sheet.Name) { Rows = sheet.Rows };
                flow.JobNames = new List<string> { sheet.Name.Split("_")[2] };
                MainFlowSheets.Add(flow);
            }
        }

        public void LoadIgxlProgramAsync(string exportFolder, List<string> usedPayloads, string jobName)
        {
            var readBinTable = new Task(() =>
            {
                BintableSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTBintablesSheet)
                    .Cast<BinTableSheet>()
                    .ToList();
            });

            var readInstance = new Task(() =>
            {
                InstanceSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTTestInstancesSheet)
                        .Cast<InstanceSheet>()
                        .ToList();
            });

            var readFlow = new Task(() =>
            {
                FlowSheetsAll =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTFlowtableSheet)
                        .Cast<FlowSheet>()
                        .ToList();

                foreach (FlowSheet flowSheet in FlowSheetsAll)
                {
                    if (flowSheet != null)
                    {
                        if (Regex.IsMatch(flowSheet.Name, "flow", RegexOptions.IgnoreCase) &&
                            Regex.IsMatch(flowSheet.Name, "main|conti|IDS|tmps|nwire", RegexOptions.IgnoreCase))
                        {
                            FlowSheets.Add(flowSheet);
                        }
                        else if (_ProcessFlowSheet(flowSheet, usedPayloads))
                        {
                            FlowSheets.Add(flowSheet);
                        }
                    }
                }
            });

            //var readFlow = new Task(() =>
            //{
            //    var files = Directory.GetFiles(exportFolder);
            //    foreach (var file in files)
            //    {
            //        IgxlSheet sheet = null;
            //        var fileName = Path.GetFileNameWithoutExtension(file);
            //        if (Regex.IsMatch(fileName, "flow", RegexOptions.IgnoreCase) &&
            //            Regex.IsMatch(fileName, "main|conti|IDS|tmps", RegexOptions.IgnoreCase))
            //            sheet = new ReadFlowSheet().GetIgxlSheet(file, SheetType.DTFlowtableSheet);
            //        else if (usedPayloads.Count > 0)// if need to reference payload with HARDIP/ATPG => read sheet
            //            sheet = new ReadFlowSheet().GetIgxlSheet(file, SheetType.DTFlowtableSheet);
            //        if (sheet != null)
            //        {
            //            if (Regex.IsMatch(sheet.Name, "flow", RegexOptions.IgnoreCase) &&
            //                Regex.IsMatch(sheet.Name, "main|conti|IDS|tmps", RegexOptions.IgnoreCase))
            //                FlowSheets.Add((FlowSheet)sheet);
            //            else if (_ProcessFlowSheet((FlowSheet)sheet, usedPayloads))
            //            {
            //                FlowSheets.Add((FlowSheet)sheet);
            //            }
            //        }
            //    }
            //});

            #region read Time/PatternSet
            var readTimeset = new Task(() =>
            {
                TimeSetSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTTimesetBasicSheet)
                        .Cast<TimeSetBasicSheet>()
                        .ToList();
            });

            var readPattern = new Task(() =>
            {
                PatSetSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTPatternSetSheet)
                        .Cast<PatSetSheet>()
                        .ToList();
            });


            var readPatSubroutine = new Task(() =>
            {
                PatSetSubSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTPatternSubroutineSheet)
                        .Cast<PatSetSubSheet>()
                        .ToList();
            });
            #endregion

            #region read DC/AC
            var readDcCategory = new Task(() =>
            {
                DcSpecSheets = IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTDCSpecSheet)
                    .Cast<DcSpecSheet>().ToList();
                DcSpecDatas = DcSpecSheets.SelectMany(p => p.Rows).ToList();
                if (DcSpecDatas.Any())
                {
                    DcCategoryList = DcSpecDatas[0].CategoryList.Select(p => p.Name).ToList();
                }
            });

            var readAcCategory = new Task(() =>
            {
                AcSpecSheets = IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTACSpecSheet)
                        .Cast<AcSpecSheet>().ToList();
            });
            #endregion

            #region read other sheets
            var readOthers = new Task(() =>
            {
                var joblist = IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTJobListSheet)
                    .Cast<JobListSheet>().ToList();
                if (joblist.Any())
                {
                    JoblistSheet = joblist.First();
                }

                {
                    JobRow userJob = JoblistSheet.Rows.FirstOrDefault(x => x.JobName.Equals(jobName, StringComparison.OrdinalIgnoreCase));
                    string jobPinMap = "";
                    if (userJob != null)
                    {
                        jobPinMap = userJob.PinMap;
                    }

                    PinMaps = IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTPinMap)
                        .Cast<PinMapSheet>().ToList();
                    PinMapSheet pinMap = null;
                    if (!string.IsNullOrEmpty(jobPinMap))
                    {
                        pinMap = PinMaps.FirstOrDefault(x => x.Name.Equals(jobPinMap, StringComparison.OrdinalIgnoreCase));
                    }

                    if (pinMap == null)
                    {
                        pinMap = PinMaps[0];
                    }

                    PinList = (List<Pin>)pinMap.PinList;
                    PinGroups = (List<PinGroup>)pinMap.GroupList;
                }

                ChannelMapSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTChanMap)
                        .Cast<ChannelMapSheet>()
                        .ToList();
                GlobalSpecRows = IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTGlobalSpecSheet)
                    .Cast<GlobalSpecSheet>().SelectMany(p => p.Rows).ToList();

                PortRows =
                    IgxlSheetReaderHelpers.GetIgxlSheets(exportFolder, EnumSheetType.DTPortMapSheet)
                        .Cast<PortMapSheet>().SelectMany(p => p.Rows).SelectMany(q => q.PortRows).ToList();
            });
            #endregion

            #region Task Start
            readDcCategory.Start();
            readAcCategory.Start();
            readBinTable.Start();
            readInstance.Start();
            readFlow.Start();
            readTimeset.Start();
            readPattern.Start();
            readPatSubroutine.Start();
            readOthers.Start();
            #endregion

            Task.WaitAll(readBinTable, readInstance, readFlow, readTimeset, readPattern, readPatSubroutine,
                readOthers, readDcCategory, readAcCategory);
        }

        public List<InstanceSheet> InstList
        {
            get
            {
                var instList = new List<InstanceSheet>();

                if (JoblistSheet == null)
                {
                    return InstanceSheets;
                }
                // get job objcet
                //var job = _testProgram.Jobs.FirstOrDefault(a => a.Name.Equals(_selectJob, StringComparison.OrdinalIgnoreCase));
                if (JoblistSheet.Rows.All
                    (p => !p.JobName.Equals(_selectJob, StringComparison.OrdinalIgnoreCase)))
                {
                    return instList;
                }

                // get inst from inst sheet specified in the job sheet
                //instList.AddRange(
                //    job.GetSheetsByType(SheetType.DTTestInstancesSheet)
                //    .Where(sh => sh != null)
                //    .Cast<TestInstanceSheet>()
                //    .SelectMany(instSheet => instSheet.TestInstances.Where(inst => inst != null  && !string.IsNullOrEmpty(inst.Name))));

                return InstanceSheets;
            }
        }

        public Dictionary<string, List<string>> ReadPatSet(string patSetFilePath)
        {
            var patSetDict = new Dictionary<string, List<string>>();

            try
            {
                // Read the file and display it line by line.  
                var file = new StreamReader(patSetFilePath);
                bool isHeader = true;
                int subPatCol = 0;
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    string[] tokens = line.Split('\t');
                    if (tokens.Length <= 3)
                    {
                        continue;
                    }

                    if (isHeader)
                    {
                        if (tokens[1] != "Pattern Set")
                        {
                            continue;
                        }

                        for (int i = 2; i < tokens.Length; i++)
                        {
                            if (tokens[i] != "File/Group Name")
                            {
                                continue;
                            }

                            subPatCol = i;
                            isHeader = false;
                            break;
                        }
                    }
                    else
                    {
                        if (tokens.Length < subPatCol)
                        {
                            continue;
                        }

                        string patName = tokens[1].ToUpper();
                        string subPat = tokens[subPatCol].ToUpper();
                        if (subPat.Contains("\\") || subPat.Contains("/"))
                        {
                            patSetDict[patName] = new List<string> { patName };
                        }
                        else if (patSetDict.ContainsKey(patName))
                        {
                            patSetDict[patName].Add(subPat);
                        }
                        else
                        {
                            patSetDict[patName] = new List<string> { subPat };
                        }
                    }
                }

                file.Close();
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message);
            }

            return patSetDict;
        }

        public List<string> GetJobList()
        {
            return JoblistSheet.Rows.Select(p => p.JobName).ToList();
        }

        public List<FlowRow> GetAllFlowSteps()
        {
            var allFlowSteps = new List<FlowRow>();
            FlowSheet mainFlowSheet = GetMainFlowSheet();
            if (mainFlowSheet == null)
            {
                return allFlowSteps;
            }

            ReadFlowSheet(mainFlowSheet, allFlowSteps);
            return allFlowSteps;
        }

        public FlowSheet GetMainFlowSheet()
        {
            JobRow usedJob = JoblistSheet?.Rows.FirstOrDefault(p =>
                p.JobName.Equals(_selectJob, StringComparison.OrdinalIgnoreCase));
            return usedJob == null
                ? null
                : FlowSheets.FirstOrDefault(p => p.Name.Equals(usedJob.FlowTable, StringComparison.OrdinalIgnoreCase));
        }

        public FlowSheet GetMainFlowSheet(JobRow job)
        {
            return FlowSheets.FirstOrDefault(p => p.Name.Equals(job.FlowTable, StringComparison.OrdinalIgnoreCase));
        }

        private void ReadFlowSheet(FlowSheet flowSheet, ICollection<FlowRow> flowSteps)
        {
            if (flowSheet == null)
            {
                return;
            }

            foreach (FlowRow flowstep in flowSheet.Rows)
            {
                if (flowstep.Opcode.ToLower().Trim() == "return")
                {
                    return;
                }

                if (flowstep.Opcode.ToLower().Trim() == "call")
                {
                    ReadFlowSheet(
                        FlowSheets.FirstOrDefault
                        (p => p.Name.Equals(flowstep.Parameter, StringComparison.OrdinalIgnoreCase))
                        , flowSteps);
                }
                else
                {
                    flowSteps.Add(flowstep);
                }
            }
        }

        private bool _ProcessFlowSheet(FlowSheet sheet, List<string> payloads)
        {
            var testItems = sheet.Rows.Where(p => p.Opcode.Equals("test", StringComparison.OrdinalIgnoreCase) &&
                                                      !string.IsNullOrEmpty(p.Parameter)).Select(s => s.Parameter).ToList();
            foreach (string item in testItems)
            {
                if (payloads.Exists(p => Regex.IsMatch(item, p, RegexOptions.IgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }

        public void LoadIgxlProgramForAutoPatAsync(string path)
        {
            var readPatSet = new Task(() =>
            {
                PatSetSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPatternSetSheet)
                        .Cast<PatSetSheet>()
                        .ToList();
            });

            var readPatSubroutine = new Task(() =>
            {
                PatSetSubSheets =
                    IgxlSheetReaderHelpers.GetIgxlSheets(path, EnumSheetType.DTPatternSubroutineSheet)
                        .Cast<PatSetSubSheet>()
                        .ToList();
            });

            readPatSet.Start();
            readPatSubroutine.Start();
            Task.WaitAll(readPatSet, readPatSubroutine);
        }
    }
}
