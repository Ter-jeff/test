using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

namespace IgxlLib
{
    public class IgxlWorkBook
    {
        public const string MainBinTblSheetName = "Bin_Table";
        public const string AcSpecsSheetName = "AC_Specs";
        public const string FlowTableMainInitEnableWd = "Flow_Table_Main_Init_EnableWd";
        public const string FlowTableMainInitFlows = "Flow_Table_Main_Init_Flows";
        public const string FlowTableMainEndFlow = "Flow_Table_Main_EndFlow";
        public const string PatSetsAll = "PatSets_All";
        public const string PatternSubroutine = "Pattern_Subroutine";

        private KeyValuePair<string, PinMapSheet> _pinMapPair;
        private KeyValuePair<string, GlobalSpecSheet> _glbSpecSheetPair;
        private readonly Dictionary<string, PSetSheet> _pSetSheets = new(StringExtensions.IgnoreCase);

        #region Property
        public Dictionary<string, SubFlowSheet> SubFlowSheets { get; } = new Dictionary<string, SubFlowSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, LimitSetsSheet> LimitSetsSheets { get; } = new Dictionary<string, LimitSetsSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, MainFlow> MainFlowSheets { get; } = new Dictionary<string, MainFlow>(StringExtensions.IgnoreCase);
        public Dictionary<string, InstanceSheet> InsSheets { get; } = new Dictionary<string, InstanceSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, DcSpecSheet> DcSpecSheets { get; } = new Dictionary<string, DcSpecSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, AcSpecSheet> AcSpecSheets { get; } = new Dictionary<string, AcSpecSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, LevelSheet> LevelSheets { get; } = new Dictionary<string, LevelSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, TimeSetBasicSheet> TimeSetSheets { get; } = new Dictionary<string, TimeSetBasicSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, PatSetSheet> PatSetSheets { get; } = new Dictionary<string, PatSetSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, PatSetSubSheet> PatSetSubSheets { get; } = new Dictionary<string, PatSetSubSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, CharSheet> CharSheets { get; } = new Dictionary<string, CharSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, JitterSheet> JitterSheets { get; } = new Dictionary<string, JitterSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, WaveDefinitionSheet> WaveDefSheets { get; } = new Dictionary<string, WaveDefinitionSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, MixedSignalSheet> MixedSignalSheets { get; } = new Dictionary<string, MixedSignalSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, BinTableSheet> BinTblSheets { get; } = new Dictionary<string, BinTableSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, ChannelMapSheet> ChannelMapSheets { get; } = new Dictionary<string, ChannelMapSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, PortMapSheet> PortMapSheets { get; } = new Dictionary<string, PortMapSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, ReferenceSheet> ReferenceSheets { get; } = new Dictionary<string, ReferenceSheet>(StringExtensions.IgnoreCase);
        public Dictionary<string, JobListSheet> JobListSheets { get; } = new Dictionary<string, JobListSheet>(StringExtensions.IgnoreCase);

        public KeyValuePair<string, GlobalSpecSheet> GlbSpecSheetPair
        {
            get { return _glbSpecSheetPair; }
            set
            {
                if (_glbSpecSheetPair.Value == null)
                {
                    string subFileSubName = Contact(value.Key, value.Value.Name);
                    _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(subFileSubName, value.Value);
                }
                else
                {
                    throw new Exception("Can't duplicate set igxl global Spec Sheet.");
                }
            }
        }
        public KeyValuePair<string, PinMapSheet> PinMapPair
        {
            get
            {
                if (_pinMapPair.Value == null)
                {
                    _pinMapPair = new KeyValuePair<string, PinMapSheet>("", new PinMapSheet(""));
                }

                return _pinMapPair;
            }
            set
            {
                if (_pinMapPair.Value == null || string.IsNullOrEmpty(_pinMapPair.Value.Name))
                {
                    string subFileSubName = Contact(value.Key, value.Value.Name);
                    _pinMapPair = new KeyValuePair<string, PinMapSheet>(subFileSubName, value.Value);
                }
            }
        }

        public Dictionary<string, string> AllIgxlSheetsDicWithType
        {
            get
            {
                return AllIgxlSheets.ToDictionary(x => x.Key, x => x.Value is MainFlow ? "Main Flow" : x.Value.IgxlSheetName);
            }
        }

        public Dictionary<string, IIgxlSheet> AllIgxlSheets
        {
            get
            {
                var allSheets = new Dictionary<string, IIgxlSheet>();
                void AddSheets<T>(IDictionary<string, T> source) where T : IIgxlSheet
                {
                    foreach (KeyValuePair<string, T> item in source)
                    {
                        allSheets[item.Key] = item.Value;
                    }
                }

                AddSheets(SubFlowSheets);
                AddSheets(LimitSetsSheets);
                AddSheets(MainFlowSheets);
                AddSheets(InsSheets);
                AddSheets(DcSpecSheets);
                AddSheets(AcSpecSheets);
                AddSheets(LevelSheets);
                AddSheets(TimeSetSheets);
                AddSheets(PatSetSheets);
                AddSheets(PatSetSubSheets);
                AddSheets(BinTblSheets);
                AddSheets(JobListSheets);
                AddSheets(ChannelMapSheets);
                AddSheets(PortMapSheets);
                AddSheets(CharSheets);
                AddSheets(WaveDefSheets);
                AddSheets(MixedSignalSheets);
                AddSheets(JitterSheets);
                AddSheets(ReferenceSheets);

                if (_pinMapPair.Key != null && _pinMapPair.Value != null)
                {
                    allSheets[_pinMapPair.Key] = _pinMapPair.Value;
                }

                if (_glbSpecSheetPair.Key != null && _glbSpecSheetPair.Value != null)
                {
                    allSheets[_glbSpecSheetPair.Key] = _glbSpecSheetPair.Value;
                }

                AddSheets(_pSetSheets);

                return allSheets;
            }
        }

        public Dictionary<string, IIgxlSheet> RestrictedIgxlSheets { get; } = [];
        #endregion

        #region Member Function
        public InstanceSheet? GetInstanceSheet(string name)
        {
            return InsSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase(name));
        }

        public AcSpecSheet? GetAcSpecsSheet(string name = AcSpecsSheetName)
        {
            return AcSpecSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase(name));
        }

        public BinTableSheet GetBinTblSheet(string path, string binTableName)
        {
            BinTableSheet? binTableSheet = BinTblSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase(binTableName));
            if (binTableSheet == null)
            {
                binTableSheet = new BinTableSheet(binTableName);
                AddBinTblSheet(path, binTableSheet);
            }
            return binTableSheet;
        }

        public BinTableSheet GetMainBinTblSheet(string path)
        {
            return GetBinTblSheet(path, MainBinTblSheetName);
        }

        public void AddSubFlowSheet(KeyValuePair<string, SubFlowSheet> sheet)
        {
            SubFlowSheets[sheet.Key] = sheet.Value;
        }

        public void AddSubFlowSheet(string fileSubPath, SubFlowSheet subFlowSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, subFlowSheet.Name);
            SubFlowSheets[subFileSubName] = subFlowSheet;
        }

        public void AddMainFlowSheet(string fileSubPath, MainFlow mainFlow)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, mainFlow.Name);
            MainFlowSheets[subFileSubName] = mainFlow;
        }

        public void AddInsSheet(string fileSubPath, InstanceSheet instanceSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, instanceSheet.Name);
            InsSheets[subFileSubName] = instanceSheet;
        }

        public void AddDcSpecSheet(string fileSubPath, DcSpecSheet dcSpecSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, dcSpecSheet.Name);
            DcSpecSheets[subFileSubName] = dcSpecSheet;
        }

        public void AddAcSpecSheet(string fileSubPath, AcSpecSheet acSpecSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, acSpecSheet.Name);
            AcSpecSheets[subFileSubName] = acSpecSheet;
        }

        public void AddLevelSheet(string fileSubPath, LevelSheet levelSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, levelSheet.Name);
            LevelSheets[subFileSubName] = levelSheet;
        }

        public void AddTimeSetSheet(string fileSubPath, TimeSetBasicSheet timeSetBasicSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, timeSetBasicSheet.Name);
            TimeSetSheets[subFileSubName] = timeSetBasicSheet;
        }

        public void AddPatSetSheet(string fileSubPath, PatSetSheet patSetSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, patSetSheet.Name);
            PatSetSheets[subFileSubName] = patSetSheet;
        }

        public void AddPatSetSubSheet(string fileSubPath, PatSetSubSheet patSetSubSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, patSetSubSheet.Name);
            PatSetSubSheets[subFileSubName] = patSetSubSheet;
        }

        public void AddBinTblSheet(string fileSubPath, BinTableSheet binTableSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, binTableSheet.Name);
            if (!BinTblSheets.TryGetValue(subFileSubName, out BinTableSheet? tblSheet))
            {
                BinTblSheets.Add(subFileSubName, binTableSheet);
            }
            else
            {
                tblSheet.Rows.AddRange(binTableSheet.Rows);
            }
        }

        public void AddJobListSheet(string fileSubPath, JobListSheet jobListSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, jobListSheet.Name);
            JobListSheets[subFileSubName] = jobListSheet;
        }

        public void AddChannelMapSheet(string fileSubPath, ChannelMapSheet channelMapSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, channelMapSheet.Name);
            ChannelMapSheets[subFileSubName] = channelMapSheet;
        }

        public void AddPortMapSheet(string fileSubPath, PortMapSheet portMapSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, portMapSheet.Name);
            PortMapSheets[subFileSubName] = portMapSheet;
        }

        public void AddWaveDefSheet(string fileSubPath, WaveDefinitionSheet waveDefinitionSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, waveDefinitionSheet.Name);
            WaveDefSheets[subFileSubName] = waveDefinitionSheet;
        }

        public void AddMixedSignalSheet(string fileSubPath, MixedSignalSheet mixedSignalSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, mixedSignalSheet.Name);
            MixedSignalSheets[subFileSubName] = mixedSignalSheet;
        }

        public void AddCharSheet(string fileSubPath, CharSheet charSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, charSheet.Name);
            CharSheets[subFileSubName] = charSheet;
        }

        public void AddJitterSheet(string fileSubPath, JitterSheet jitterSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, jitterSheet.Name);
            JitterSheets[subFileSubName] = jitterSheet;
        }

        public void AddReferenceSheet(string fileSubPath, ReferenceSheet referenceSheet)
        {
            string subFileSubName = GetSubFileSubName(fileSubPath, referenceSheet.Name);
            ReferenceSheets[subFileSubName] = referenceSheet;
        }

        public List<string> GetFlowSheetNameList()
        {
            var flowSheetNameList = new List<string>();
            foreach (KeyValuePair<string, SubFlowSheet> flowsheet in SubFlowSheets)
            {
                flowSheetNameList.Add(flowsheet.Value.Name);
            }

            return flowSheetNameList;
        }

        public static string GetSubFileSubName(string fileSubPath, string sheetName)
        {
            return Contact(fileSubPath, sheetName);
        }

        public void PrintAllSheets(string igxlVersion = "10.10")
        {
            Parallel.ForEach(AllIgxlSheets, allIgxlSheet =>
            {
                string sheetVersion = IgxlLoaderHelpers.GetSheetVersion(allIgxlSheet.Value, igxlVersion);
                allIgxlSheet.Value.Write(allIgxlSheet.Key.Replace('\\', Path.DirectorySeparatorChar) + ".txt", sheetVersion);
            });
            Parallel.ForEach(RestrictedIgxlSheets, restrictedIgxlSheet =>
            {
                string sheetVersion = IgxlLoaderHelpers.GetSheetVersion(restrictedIgxlSheet.Value, igxlVersion);
                restrictedIgxlSheet.Value.Write(restrictedIgxlSheet.Key.Replace('\\', Path.DirectorySeparatorChar) + ".txt", sheetVersion);
            });
        }

        public void Clear()
        {
            SubFlowSheets.Clear();
            LimitSetsSheets.Clear();
            MainFlowSheets.Clear();
            InsSheets.Clear();
            DcSpecSheets.Clear();
            AcSpecSheets.Clear();
            LevelSheets.Clear();
            TimeSetSheets.Clear();
            PatSetSheets.Clear();
            PatSetSubSheets.Clear();
            BinTblSheets.Clear();
            JobListSheets.Clear();
            ChannelMapSheets.Clear();
            CharSheets.Clear();
            WaveDefSheets.Clear();
            MixedSignalSheets.Clear();
            _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(null!, null!);
            _pinMapPair = new KeyValuePair<string, PinMapSheet>(null!, null!);
            PortMapSheets.Clear();
            ReferenceSheets.Clear();
            JitterSheets.Clear();
            RestrictedIgxlSheets.Clear();
        }

        public static string Contact(string mainDir, string subDir)
        {
            if (mainDir != null && subDir != null)
            {
                char sep = Path.DirectorySeparatorChar;
                mainDir = mainDir.Replace('\\', sep).Replace('/', sep).TrimEnd(sep);
                subDir = subDir.Replace('\\', sep).Replace('/', sep).TrimStart(sep);
                return mainDir + sep + subDir;
            }
            throw new Exception($"Directory object: {mainDir} or {subDir} is null. ");
        }

        public void GenAllFailFlagToMainInitEnableWd(IEnumerable<string> flags, string columnA, string dirMain, bool merge = true)
        {
            SubFlowSheet mainInitEnableWdSheet = GetFlowSheet(FlowTableMainInitEnableWd, dirMain, "Initialize");
            if (mainInitEnableWdSheet == null)
            {
                return;
            }

            List<string> allInitFalse = mainInitEnableWdSheet.GetAllInitFalse();
            flags = [.. flags.Except(allInitFalse, StringExtensions.IgnoreCase).Where(x => !string.IsNullOrEmpty(x))];
            if (flags.Any())
            {
                var flowRows = new List<FlowRow>();
                if (merge)
                {
                    var flagList = new List<string>();
                    string opcode = !string.IsNullOrEmpty(columnA) && (columnA.EqualsIgnoreCase("Efuse") || columnA.StartsWithIgnoreCase("AteStrSummary") || columnA.EqualsIgnoreCase("HarvestInit"))
                                ? OpCode.FlagClear
                                : OpCode.FlagFalse;
                    FlowRow? existRow = mainInitEnableWdSheet.Rows.FindLast(x => x.ColumnA == columnA && x.Opcode == opcode);
                    bool isExist = false;
                    var allFlags = new List<string>();
                    if (existRow != null)
                    {
                        List<string> existFlags = [.. existRow.Parameter.Split(',')];
                        allFlags.AddRange(existFlags);
                        isExist = true;
                    }
                    allFlags.AddRange(flags);
                    foreach (string flag in allFlags)
                    {
                        flagList.Add(flag);
                        if (string.Join(",", flagList).Length >= 6000)
                        {
                            if (isExist)
                            {
                                existRow!.Parameter = string.Join(",", flagList);
                                isExist = false;
                            }
                            else
                            {
                                var flowRow = new FlowRow
                                {
                                    Opcode = opcode,
                                    Parameter = string.Join(",", flagList),
                                    ColumnA = columnA
                                };
                                mainInitEnableWdSheet.Rows.Add(flowRow);
                            }
                            flagList = [];
                        }
                    }
                    if (flagList.Count != 0)
                    {
                        if (isExist)
                        {
                            existRow!.Parameter = string.Join(",", flagList);
                        }
                        else
                        {
                            var flowRow = new FlowRow
                            {
                                Opcode = opcode,
                                Parameter = string.Join(",", flagList),
                                ColumnA = columnA
                            };
                            mainInitEnableWdSheet.Rows.Add(flowRow);
                        }
                    }
                }
                else
                {
                    foreach (string failFlag in flags)
                    {
                        var flowRow = new FlowRow
                        {
                            Opcode = !string.IsNullOrEmpty(columnA) &&
                                     (columnA.EqualsIgnoreCase("Efuse") || columnA.StartsWithIgnoreCase("AteStrSummary") || columnA.EqualsIgnoreCase("HarvestInit"))
                                ? OpCode.FlagClear
                                : OpCode.FlagFalse,
                            Parameter = failFlag,
                            ColumnA = columnA
                        };
                        if (!mainInitEnableWdSheet.Rows.Exists(x => x.Parameter.EqualsIgnoreCase(flowRow.Parameter)))
                        {
                            flowRows.Add(flowRow);
                        }
                    }
                    mainInitEnableWdSheet.AddRows(flowRows);
                }
            }
        }

        public void GenEnableToMainInitEnableWd(string enable, string op, string columnA, string dirMain, string? parameter = null)
        {
            SubFlowSheet mainInitEnableWdSheet = GetFlowSheet(FlowTableMainInitEnableWd, dirMain, "Initialize");
            if (mainInitEnableWdSheet == null)
            {
                return;
            }

            var flowRow = new FlowRow { Enable = enable, Opcode = op, ColumnA = columnA };
            if (!string.IsNullOrEmpty(parameter))
            {
                flowRow.Parameter = parameter;
            }

            mainInitEnableWdSheet.AddRow(flowRow);
        }

        public SubFlowSheet GetFlowSheet(string sheetName, string dir, string source = "")
        {
            SubFlowSheet? existingFlowSheet = SubFlowSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase(sheetName));
            if (existingFlowSheet != null)
            {
                return existingFlowSheet;
            }

            var flowSheet = new SubFlowSheet(sheetName, source);
            AddSubFlowSheet(dir, flowSheet);
            return flowSheet;
        }

        public void SetFlowSheet(string flowName, SubFlowSheet subFlowSheet)
        {
            foreach (KeyValuePair<string, SubFlowSheet> sheet in SubFlowSheets)
            {
                if (sheet.Value.Name.EqualsIgnoreCase(flowName))
                {
                    SubFlowSheets[sheet.Key] = subFlowSheet;
                    break;
                }
            }
        }

        public PatSetSheet GetPatSetsSheet(string sheetName, string dir)
        {
            PatSetSheet? existingPatSetSheet = PatSetSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase(sheetName));
            if (existingPatSetSheet != null)
            {
                return existingPatSetSheet;
            }

            var patSetSheet = new PatSetSheet(sheetName);
            AddPatSetSheet(dir, patSetSheet);
            return patSetSheet;
        }

        public PatSetSubSheet GetPatSetSubSheet(string sheetName, string dir)
        {
            PatSetSubSheet? existingPatSetSubSheet = PatSetSubSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase(sheetName));
            if (existingPatSetSubSheet != null)
            {
                return existingPatSetSubSheet;
            }

            var patSetSubSheet = new PatSetSubSheet(sheetName);
            AddPatSetSubSheet(dir, patSetSubSheet);
            return patSetSubSheet;
        }

        public bool TryGetTestInstCommon(out InstanceSheet? instanceSheet)
        {
            instanceSheet = InsSheets.Values.FirstOrDefault(x => x.Name.EqualsIgnoreCase("TestInst_Common"));
            return instanceSheet != null;
        }

        public List<FlowRow> GetSubFlows(FlowRow flowRow, string? opCodeFilter = null)
        {
            // Pre-allocate capacity to reduce resizing cost
            var result = new List<FlowRow>(4096);

            // Entry must be a Call node; include it as the root of the subflow search
            if (!flowRow.Opcode.EqualsIgnoreCase(OpCode.Call) ||
                string.IsNullOrEmpty(flowRow.Parameter))
            {
                return result;
            }
            result.Add(flowRow);

            // Build sheet lookup dictionary once (case-insensitive)
            var dic = SubFlowSheets.Values
                .GroupBy(x => x.Name, StringExtensions.IgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringExtensions.IgnoreCase);

            // Recursion stack to prevent cycle in current path
            var stack = new HashSet<string>(StringExtensions.IgnoreCase);

            // Cache flattened rows per sheet (unfiltered)
            var cache = new Dictionary<string, List<FlowRow>>(StringExtensions.IgnoreCase);

            void Collect(string sheetName)
            {
                // Prevent cycle (only current path)
                if (!stack.Add(sheetName))
                {
                    return;
                }

                // Use cache if available
                if (cache.TryGetValue(sheetName, out List<FlowRow>? cachedRows))
                {
                    if (string.IsNullOrEmpty(opCodeFilter))
                    {
                        result.AddRange(cachedRows);
                    }
                    else
                    {
                        foreach (FlowRow r in cachedRows)
                        {
                            if (r.Opcode.EqualsIgnoreCase(opCodeFilter))
                            {
                                result.Add(r);
                            }
                        }
                    }

                    stack.Remove(sheetName);
                    return;
                }

                // If sheet not found, exit
                if (!dic.TryGetValue(sheetName, out SubFlowSheet? sheet))
                {
                    stack.Remove(sheetName);
                    return;
                }

                var local = new List<FlowRow>(128);

                // Inline traversal to preserve order
                foreach (FlowRow row in sheet.Rows)
                {
                    if (row.IsBackup)
                    {
                        continue;
                    }

                    if (row.Opcode.EqualsIgnoreCase(OpCode.Call))
                    {
                        Collect(row.Parameter);
                        continue;
                    }

                    // Store in cache
                    local.Add(row);

                    // Apply optional filter
                    if (string.IsNullOrEmpty(opCodeFilter) || row.Opcode.EqualsIgnoreCase(opCodeFilter))
                    {
                        result.Add(row);
                    }
                }

                // Save cache (unfiltered)
                cache[sheetName] = local;

                stack.Remove(sheetName);
            }

            Collect(flowRow.Parameter);

            return result;
        }
        #endregion
    }
}
