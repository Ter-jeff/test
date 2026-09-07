using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.Utility;

namespace TestPlanLib.Harvest
{
    public class HarvestingTruthTableReader(string truthTableSheetName) : MySheetReader<HarvestingTruthTableSheet>
    {
        private const string ConHeaderReadFuse = "Read Fuses";
        private const string ConHeaderFlagResult = "Flag Result";
        private const string ConHeaderFlagResult1 = "Test Flow Flag Result";
        private const string ConHeaderHarvestDecision = "Harvest Decision";
        private const string ConHeaderOutputFlags = "Output Flags";
        private const string ConHeaderHarvestDecisionFlag = "Harvest Decision Output Flag";
        private const string ConHeaderFusing = "FUSING";
        private const string ConHeaderProposedBinName = "proposed bin name";
        private const string ConHeaderHardBin = "Hard Bin";
        private const string ConHeaderInputFlags = @"INPUT Flags ->";
        private const string ConHeaderRepairFlag = "Repair Flag";

        private string _sheetName = "";
        private readonly HeaderInfo _indexReadFuse = new();
        private readonly HeaderInfo _indexDomain = new();
        private readonly HeaderInfo _indexHarvestDecision = new();
        private readonly HeaderInfo _indexOutputFlags = new();
        private readonly HeaderInfo _indexOutputHarvestDecisionFlags = new();
        private readonly HeaderInfo _indexFusing = new();
        private readonly HeaderInfo _indexProposedBinName = new();
        private readonly HeaderInfo _indexHardBin = new();
        private readonly HeaderInfo _indexInputFlag = new();
        private readonly HeaderInfo _indexRepairFlag = new();
        private readonly string _commentTruthTableSheetName = truthTableSheetName;

        public override HarvestingTruthTableSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            var harvestingTruthTableSheet = new HarvestingTruthTableSheet(_sheetName);

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return harvestingTruthTableSheet;
            }

            GetHeaderIndex();

            harvestingTruthTableSheet = ReadSheet(_sheetName);

            return harvestingTruthTableSheet;
        }

        private HarvestingTruthTableSheet ReadSheet(string sheetName)
        {
            var harvestingTruthTableSheet = new HarvestingTruthTableSheet(sheetName)
            {
                Job = HarvestingTruthTableReaderHelpers.GetJob(sheetName, _commentTruthTableSheetName),
                StartRowNum = StartRow
            };

            ReadData(sheetName, harvestingTruthTableSheet, _indexInputFlag);

            harvestingTruthTableSheet.IndexDomain = _indexDomain.Start;
            harvestingTruthTableSheet.IndexHarvestDecision = _indexHarvestDecision.Start;
            harvestingTruthTableSheet.IndexFusing = _indexFusing.Start;
            harvestingTruthTableSheet.IndexReadFuse = _indexReadFuse.Start;
            harvestingTruthTableSheet.IndexProposedBinName = _indexProposedBinName.Start;
            harvestingTruthTableSheet.IndexHardBin = _indexHardBin.Start;

            return harvestingTruthTableSheet;
        }

        private void ReadData(string sheetName, HarvestingTruthTableSheet harvestingTruthTableSheet, HeaderInfo headerInfo)
        {
            int headerRowIndex = headerInfo.Start;
            int dataStartRowIndex = headerRowIndex + 1;
            var sumFlagsDic = new Dictionary<string, List<string>>();
            for (int i = dataStartRowIndex; i <= EndRow; i++)
            {
                var row = new HarvestingTruthTableRow(sheetName) { RowNum = i };
                if (_indexDomain.Start != -1)
                {
                    if (_indexReadFuse.Start - 1 >= 1)
                    {
                        row.Condition = ExcelWorksheet.GetCellValue(i, _indexReadFuse.Start - 1).Trim();
                        if (string.IsNullOrEmpty(row.Condition))
                        {
                            continue;
                        }
                    }
                    else if (_indexDomain.Start - 1 >= 1)
                    {
                        row.Condition = ExcelWorksheet.GetCellValue(i, _indexDomain.Start - 1).Trim();
                        if (string.IsNullOrEmpty(row.Condition))
                        {
                            continue;
                        }
                    }

                    HarvestingTruthTableReaderHelpers.AddFlags(ExcelWorksheet, _indexDomain, sheetName, harvestingTruthTableSheet, headerRowIndex, sumFlagsDic, i, row);
                }

                SetRow(ExcelWorksheet, headerRowIndex, i, row);

                if (_indexProposedBinName.Start != -1)
                {
                    row.ProposedBinName = ExcelWorksheet.GetCellValue(i, _indexProposedBinName.Start).Trim();
                    if (string.IsNullOrEmpty(row.ProposedBinName))
                    {
                        continue;
                    }
                }

                if (_indexHardBin.Start != -1)
                {
                    row.HardBin = ExcelWorksheet.GetCellValue(i, _indexHardBin.Start).Trim();
                    if (string.IsNullOrEmpty(row.HardBin))
                    {
                        continue;
                    }
                }

                if (_indexReadFuse.Start != -1)
                {
                    row.ReadFuse = HarvestingTruthTableReaderHelpers.GetReadFuses(ExcelWorksheet, _indexReadFuse, headerRowIndex, i);
                }

                if (_indexRepairFlag.Start != -1)
                {
                    row.RepairFlags = HarvestingTruthTableReaderHelpers.GetRepairFlags(ExcelWorksheet, _indexRepairFlag, headerRowIndex, i);
                }

                harvestingTruthTableSheet.Rows.Add(row);
            }

            foreach (Flag flag in harvestingTruthTableSheet.Rows.SelectMany(row => row.Flags).Where(x => x.IsSum))
            {
                if (sumFlagsDic.TryGetValue(flag.FlagName, out List<string>? value))
                {
                    flag.SumFlags.AddRange(value);
                }
            }
        }

        private void SetRow(ExcelWorksheet excelWorksheet, int headerRowIndex, int i, HarvestingTruthTableRow harvestingTruthTableRow)
        {
            if (_indexHarvestDecision.Start != -1)
            {
                for (int j = _indexHarvestDecision.Start; j <= _indexHarvestDecision.End; j++)
                {
                    string outputFlag = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                    if (!harvestingTruthTableRow.OutputBinTables.ContainsKey(outputFlag) && !string.IsNullOrEmpty(outputFlag))
                    {
                        harvestingTruthTableRow.OutputBinTables.Add(outputFlag, excelWorksheet.GetCellValue(i, j).Trim());
                    }
                }
            }

            if (_indexOutputHarvestDecisionFlags.Start != -1)
            {
                for (int j = _indexOutputHarvestDecisionFlags.Start; j <= _indexOutputHarvestDecisionFlags.End; j++)
                {
                    string outputFlag = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                    if (!harvestingTruthTableRow.OutputHarvestDecisionFlags.ContainsKey(outputFlag) && !string.IsNullOrEmpty(outputFlag))
                    {
                        harvestingTruthTableRow.OutputHarvestDecisionFlags.Add(outputFlag, excelWorksheet.GetCellValue(i, j).Trim());
                    }
                }
            }

            if (_indexOutputFlags.Start != -1)
            {
                for (int j = _indexOutputFlags.Start; j <= _indexOutputFlags.End; j++)
                {
                    string outputFlag = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                    if (!harvestingTruthTableRow.OutputFlags.ContainsKey(outputFlag) && !string.IsNullOrEmpty(outputFlag))
                    {
                        harvestingTruthTableRow.OutputFlags.Add(outputFlag, excelWorksheet.GetCellValue(i, j).Trim());
                    }
                }
            }

            if (_indexFusing.Start != -1)
            {
                for (int j = _indexFusing.Start; j <= _indexFusing.End; j++)
                {
                    string field = excelWorksheet.GetCellValue(headerRowIndex - 1, j).Trim();
                    string address = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                    bool isSingleAddress = string.IsNullOrEmpty(excelWorksheet.MergedCells[headerRowIndex, j]);
                    var fusing = new Fusing { Field = field, Address = address };

                    string value = excelWorksheet.GetCellValue(i, j).Trim();
                    GroupItem iterates = value.GetIterates();
                    string groupValue = HarvestingTruthTableReaderHelpers.GetMergedCellValueWithGroup(excelWorksheet, i, j, value, iterates.Groups).Trim();
                    fusing.Value = HarvestingTruthTableReaderHelpers.GetfusingValue(address, isSingleAddress, value, iterates, groupValue);

                    if (!string.IsNullOrEmpty(address))
                    {
                        harvestingTruthTableRow.Fusings.Add(fusing);
                    }
                }
            }
        }

        private void GetHeaderIndex()
        {
            var headerInfo = new Dictionary<string, List<int>>();
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();

                if (header.Length == 0)
                {
                    continue;
                }

                if (headerInfo.TryGetValue(header, out List<int>? value))
                {
                    value.Add(i);
                }
                else
                {
                    headerInfo.Add(header, [i]);
                }
            }
            for (int i = StartRow; i <= EndRow; i++)
            {
                string header = ExcelWorksheet.GetCellValue(i, StartCol).Trim();
                if (header.EqualsIgnoreCase(ConHeaderInputFlags))
                {
                    _indexInputFlag.Start = i;
                    _indexInputFlag.End = i;
                    break;
                }
            }
            foreach (KeyValuePair<string, List<int>> header in headerInfo)
            {
                if (header.Key.EqualsIgnoreCase(ConHeaderReadFuse))
                {
                    _indexReadFuse.Start = header.Value.Min();
                    _indexReadFuse.End = header.Value.Max();
                    continue;
                }
                if (header.Key.EqualsIgnoreCase(ConHeaderFlagResult) ||
                    header.Key.EqualsIgnoreCase(ConHeaderFlagResult1))
                {
                    _indexDomain.Start = header.Value.Min();
                    _indexDomain.End = header.Value.Max();
                    continue;
                }

                if (header.Key.EqualsIgnoreCase(ConHeaderHarvestDecision))
                {
                    _indexHarvestDecision.Start = header.Value.Min();
                    _indexHarvestDecision.End = header.Value.Max();
                    continue;
                }
                if (header.Key.EqualsIgnoreCase(ConHeaderOutputFlags))
                {
                    _indexOutputFlags.Start = header.Value.Min();
                    _indexOutputFlags.End = header.Value.Max();
                    continue;
                }
                if (header.Key.EqualsIgnoreCase(ConHeaderHarvestDecisionFlag))
                {
                    _indexOutputHarvestDecisionFlags.Start = header.Value.Min();
                    _indexOutputHarvestDecisionFlags.End = header.Value.Max();
                    continue;
                }
                if (header.Key.StartsWithIgnoreCase(ConHeaderFusing))
                {
                    _indexFusing.Start = header.Value.Min();
                    _indexFusing.End = header.Value.Max();
                }
                if (header.Key.StartsWithIgnoreCase(ConHeaderProposedBinName))
                {
                    _indexProposedBinName.Start = header.Value.Min();
                    _indexProposedBinName.End = header.Value.Max();
                }
                if (header.Key.StartsWithIgnoreCase(ConHeaderHardBin))
                {
                    _indexHardBin.Start = header.Value.Min();
                    _indexHardBin.End = header.Value.Max();
                }
                if (header.Key.StartsWithIgnoreCase(ConHeaderRepairFlag))
                {
                    _indexRepairFlag.Start = header.Value.Min();
                    _indexRepairFlag.End = header.Value.Max();
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderFlagResult) ||
                        ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderFlagResult1))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            ErrorMessageBox.Show($"Cannot found \"Flag Result\" header in {_sheetName}, plz check!!!");
            return false;
        }
    }

    public partial class HarvestingTruthTableSheet : MySheet
    {
        [GeneratedRegex(@"\d+", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\d+\-\d+", RegexOptions.Compiled)]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"[^\w]", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();

        public List<HarvestingTruthTableRow> Rows { set; get; }
        public string Job { set; get; } = "";
        public bool HasFuseRead
        {
            get
            {
                return Rows.Any(x => x.ReadFuse.Count != 0);
            }
        }
        public bool HasFuseWrite
        {
            get
            {
                return Rows.Any(x => x.Fusings.Count != 0);
            }
        }
        public List<string> Stage
        {
            get
            {
                var stage = new List<string>();
                if (HasFuseRead)
                {
                    stage.Add($"{Job}_RD");
                }

                stage.Add($"{Job}_OB");
                if (HasFuseWrite)
                {
                    stage.Add($"{Job}_BC");
                }

                return stage;
            }
        }
        public int IndexDomain = -1;
        public int IndexHarvestDecision = -1;
        public int IndexFusing = -1;
        public int IndexReadFuse = -1;
        public int IndexProposedBinName = -1;
        public int IndexHardBin = -1;
        public int StartRowNum = -1;
        public int SetBinFuseBit = -1;
        public int ReadBinFuseBit = -1;

        public Dictionary<string, List<string>> MergeValueByFlag = [];
        public List<string> ReadFuseChkList = [];

        #region Constructor
        public HarvestingTruthTableSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }

        public HarvestingTruthTableSheet(HarvestingTruthTableSheet harvestingTruthTableSheet) : base(harvestingTruthTableSheet)
        {
            if (harvestingTruthTableSheet == null)
            {
                Rows = [];
                return;
            }

            Rows = [.. harvestingTruthTableSheet.Rows.Select(x => x.Copy())];
            Job = harvestingTruthTableSheet.Job;
            IndexDomain = harvestingTruthTableSheet.IndexDomain;
            IndexHarvestDecision = harvestingTruthTableSheet.IndexHarvestDecision;
            IndexFusing = harvestingTruthTableSheet.IndexFusing;
            IndexReadFuse = harvestingTruthTableSheet.IndexReadFuse;
            IndexProposedBinName = harvestingTruthTableSheet.IndexProposedBinName;
            IndexHardBin = harvestingTruthTableSheet.IndexHardBin;
            StartRowNum = harvestingTruthTableSheet.StartRowNum;
            SetBinFuseBit = harvestingTruthTableSheet.SetBinFuseBit;
            ReadBinFuseBit = harvestingTruthTableSheet.ReadBinFuseBit;
            MergeValueByFlag = harvestingTruthTableSheet.MergeValueByFlag.ToDictionary(x => x.Key, x => x.Value.ToList());
            ReadFuseChkList = [.. harvestingTruthTableSheet.ReadFuseChkList];
        }

        public HarvestingTruthTableSheet Copy()
        {
            return new HarvestingTruthTableSheet(this);
        }
        #endregion

        public List<string> GetAllHarvFlagList()
        {
            return [.. Rows.First().Flags.SelectMany(x => x.AllFlags)];
        }

        public bool Check()
        {
            if (CheckJob() || CheckHeader() || CheckOrder() || CheckRows())
            {
                ErrorMessageBox.Show($"Found Harvest truth table error in {SheetName}, Plz Check!!!");
                return true;
            }
            return false;
        }

        public void CheckEfuseRead()
        {
            if (ReadFuseChkList.Count == 0)
            {
                return;
            }

            foreach (string flag in ReadFuseChkList)
            {
                if (!GetHarvestEfuseRead().ToList().Exists(x => x.EqualsIgnoreCase(flag)))
                {
                    AddError(HarvestErrorType.E_MissingFlag_01, SheetName, -1, -1, [flag]);
                }
            }
            foreach (string flag in GetHarvestEfuseRead())
            {
                if (!ReadFuseChkList.Exists(x => x.EqualsIgnoreCase(flag)))
                {
                    AddError(HarvestErrorType.E_RedundantFlag_01, SheetName, -1, -1, [flag]);
                }
            }
        }

        private bool CheckJob()
        {
            if (Job?.Length == 0)
            {
                //string errorMessage = "Job is not defined in sheet name";
                AddError(HarvestErrorType.E_InvalidFormat_01, SheetName, -1, -1, [SheetName]);
                return true;
            }
            return false;
        }

        private bool CheckRows()
        {
            foreach (HarvestingTruthTableRow row in Rows)
            {
                //Check input flag
                for (int i = 0; i < row.Flags.Count; i++)
                {
                    Flag flag = row.Flags.ElementAt(i);
                    if (!(flag.Value.EqualsIgnoreCase("X") || _regex.IsMatch(flag.Value) || _regex2.IsMatch(flag.Value)))
                    {
                        AddError(HarvestErrorType.E_InvalidFormat_02, SheetName, row.RowNum, IndexDomain + i, [flag.Value]);
                    }
                }

                //Check condition
                if (_regex3.IsMatch(row.Condition))
                {
                    AddError(HarvestErrorType.W_InvalidFormat_01, SheetName, row.RowNum, IndexDomain - 1, [row.Condition]);
                }
            }
            return false;
        }

        private bool CheckHeader()
        {
            if (IndexDomain == -1)
            {
                AddError(HarvestErrorType.E_MissingHeader_01, SheetName, -1, -1, ["\"Flag Result\""]);
                return true;
            }
            if (IndexReadFuse == -1 && !SheetName.Contains("CP1"))
            {
                AddError(HarvestErrorType.W_MissingHeader_01, SheetName, -1, -1, ["\"Read Fuses\""]);
            }
            if (IndexHarvestDecision == -1)
            {
                AddError(HarvestErrorType.E_MissingHeader_01, SheetName, -1, -1, ["\"Harvest Decision\""]);
                return true;
            }
            if (IndexFusing == -1)
            {
                AddError(HarvestErrorType.W_MissingHeader_02, SheetName, -1, -1, [Job]);
            }
            if (IndexProposedBinName == -1)
            {
                AddError(HarvestErrorType.W_MissingHeader_01, SheetName, -1, -1, ["\"proposed bin name\""]);
            }
            return false;
        }

        private bool CheckOrder()
        {
            if (IndexReadFuse != -1 && IndexDomain != -1 && IndexReadFuse > IndexDomain)
            {
                AddError(HarvestErrorType.E_InvalidOrder_01, SheetName, StartRowNum, IndexDomain, ["\"Flag Result\"", "\"Read Fuses\""]);
                return true;
            }
            if (IndexDomain != -1 && IndexHarvestDecision != -1 && IndexDomain > IndexHarvestDecision)
            {
                AddError(HarvestErrorType.E_InvalidOrder_01, SheetName, StartRowNum, IndexHarvestDecision, ["\"Harvest Decision\"", $"\"FUSING({Job})\""]);
                return true;
            }
            if (IndexHarvestDecision != -1 && IndexFusing != -1 && IndexHarvestDecision > IndexFusing)
            {
                AddError(HarvestErrorType.E_InvalidOrder_01, SheetName, StartRowNum, IndexFusing, [$"\"FUSING({Job})\"", "\"Harvest Decision\""]);
                return true;
            }
            if (IndexFusing != -1 && IndexProposedBinName != -1 && IndexFusing > IndexProposedBinName)
            {
                AddError(HarvestErrorType.E_InvalidOrder_01, SheetName, StartRowNum, IndexProposedBinName, ["\"proposed bin name\"", "\"FUSING\""]);
                return true;
            }
            return false;
        }

        public string[] GetHarvestEfuseRead()
        {
            string[] fusings = new string[Rows.Max(x => x.ReadFuse.Count)];

            foreach (HarvestingTruthTableRow row in Rows)
            {
                for (int i = 0; i < row.ReadFuse.Count; i++)
                {
                    Fusing fusing = row.ReadFuse.ElementAt(i);
                    if (fusing.Address.Contains("=>"))
                    {
                        fusings[i] = fusing.Address.Split('>').Last().Replace(" ", "");
                    }
                    else if (fusing.Value.StartsWithIgnoreCase("F_"))
                    {
                        fusings[i] = fusing.Value;
                    }
                    else
                    {
                        fusings[i] = "";
                    }
                }
            }

            return [.. fusings.Where(p => p != null)];
        }

        public List<Flag> GetDistinctFlags(List<string> failFlags)
        {
            var flagNames = new List<Flag>();
            List<Flag> flags = Rows.First().Flags;
            foreach (string failFlag in failFlags)
            {
                Flag? matchedFlag = flags.Find(x => x.FlagName.EqualsIgnoreCase(failFlag));
                if (matchedFlag != null)
                {
                    flagNames.Add(matchedFlag);
                }

                foreach (Flag flag in flags)
                {
                    if (flag.AllFlags.Exists(x => x.EqualsIgnoreCase(failFlag)))
                    {
                        flagNames.Add(flag);
                    }
                }
            }
            return [.. flagNames.Distinct()];
        }

        public string HarvFailCoreSumFlag()
        {
            List<List<string>> flags = GetHarvFailCoreSumFlags();
            return string.Join("|", flags.Select(x => string.Join("+", x)));
        }

        private List<List<string>> GetHarvFailCoreSumFlags()
        {
            var flags = new List<List<string>>();
            foreach (Flag flag in Rows.First().Flags)
            {
                if (flag.IsSum)
                {
                    List<string> passNumbers = MergeValueByFlag[flag.FlagName];
                    var harvAllCorePassFlags = passNumbers.Select(x => flag.FlagName + "_" + x).ToList();
                    foreach (string anotherFlag in GetAnotherSumFlags(flag))
                    {
                        if (!harvAllCorePassFlags.Contains(anotherFlag))
                        {
                            harvAllCorePassFlags.Add(anotherFlag);
                        }
                    }
                    flags.Add(harvAllCorePassFlags);
                }
            }

            return flags;
        }

        public List<string> HarvFailCoreSumAllFlags()
        {
            return [.. GetHarvFailCoreSumFlags().SelectMany(x => x)];
        }

        public string SumFailCoreFlags()
        {
            var flags = new List<List<string>>();
            foreach (Flag flag in Rows.First().Flags)
            {
                if (flag.IsSum)
                {
                    List<string> passNumbers = MergeValueByFlag[flag.FlagName];
                    var harvAllCorePassFlags = passNumbers.Select(x => flag.FlagName + "_" + x).ToList();
                    foreach (string anotherFlag in GetAnotherSumFlags(flag))
                    {
                        if (!harvAllCorePassFlags.Contains(anotherFlag))
                        {
                            harvAllCorePassFlags.Add(anotherFlag);
                        }
                    }
                    if (!harvAllCorePassFlags.Any(x => x.EqualsIgnoreCase(flag.FlagName + "_0")))
                    {
                        harvAllCorePassFlags.Add(flag.FlagName + "_0");
                    }

                    flags.Add(harvAllCorePassFlags);
                }
            }
            return string.Join("|", flags.Select(x => string.Join("+", x)));
        }

        private List<string> GetAnotherSumFlags(Flag flag)
        {
            List<string> passNumbers = MergeValueByFlag[flag.FlagName];
            var values = passNumbers.ToList();
            int max = 0;
            int min = 0;
            var otherSumFlags = new List<string>();
            foreach (string value in values)
            {
                if (!value.Contains("To"))
                {
                    continue;
                }

                _ = int.TryParse(value.Replace("To", "-").Split('-').First(), out int from);
                _ = int.TryParse(value.Replace("To", "-").Split('-').Last(), out int to);
                min = from < min ? from : min;
                max = to > max ? to : max;
            }
            for (int i = min; i <= max; i++)
            {
                string anotherFlagName = flag.FlagName + "_" + i;
                otherSumFlags.Add(anotherFlagName);
            }
            return otherSumFlags;
        }

        public string CustHarvFailCoreSumFlag()
        {
            var flags = new List<List<string>>();
            foreach (Flag flag in Rows.First().Flags)
            {
                if (flag.IsCust)
                {
                    List<string> passNumbers = MergeValueByFlag[flag.FlagName];
                    var harvAllCorePassFlags = passNumbers.Select(x => flag.FlagName + "_" + x).ToList();
                    flags.Add(harvAllCorePassFlags);
                }
            }
            return string.Join("|", flags.Select(x => string.Join("+", x)));
        }

        public List<string> GetHarvAllFlagList()
        {
            HarvestingTruthTableRow row1 = Rows.First();
            var flags = row1.Flags.SelectMany(x => x.AllFlags).ToList();
            flags.AddRange(row1.Flags.SelectMany(x => x.SumFlags));
            flags.AddRange(row1.OutputBinTables.Where(x => x.Value?.Length != 0).Select(x => x.Key));
            flags.AddRange(HarvFailCoreSumFlag().Split('+', '|'));
            flags.AddRange(CustHarvFailCoreSumFlag().Split('+', '|'));
            flags.AddRange(row1.HarvBinOutFailFlag().Split('+', '|'));
            return flags;
        }

        public List<string> GetRepairFlagList()
        {
            HarvestingTruthTableRow row1 = Rows.First();
            var flags = row1.RepairFlags.Select(x => x.FlagName).Distinct().ToList();
            return flags;
        }
    }

    public partial class Flag
    {
        private const string Regex1 = @"^Sum\((?<flag>.+)\)";
        private const string Regex2 = @"(?<cores>\[.+\])";

        [GeneratedRegex("^F_", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        [GeneratedRegex(Regex1, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(Regex2, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex4 = MyRegex2();

        public string Syntax
        {
            get
            {
                return Name.Split('>').Last().Replace(" ", "");
            }
        }
        public string FlagWithoutCore
        {
            get
            {
                string flag = _regex2.Match(Name).Groups["flag"].ToString();
                string flagRemoveCode = _regex4.Replace(flag, "");
                return flagRemoveCode;
            }
        }
        public string Name { set; get; } = "";
        public string FlagName { set; get; } = "";
        public string Value { set; get; } = "";
        public bool IsSum { set; get; }
        public bool IsCust { set; get; }
        public List<string> SumFlags { set; get; } = [];
        public List<string> OrFlags { set; get; } = [];
        public List<string> AllFlags { set; get; } = [];
        public bool IsMultiGroup { set; get; }
        public string EnableStage { set; get; } = "";

        public Flag() { }

        public Flag(Flag flag)
        {
            if (flag == null)
            {
                return;
            }

            Name = flag.Name;
            FlagName = flag.FlagName;
            Value = flag.Value;
            IsSum = flag.IsSum;
            IsCust = flag.IsCust;
            SumFlags = [.. flag.SumFlags];
            OrFlags = [.. flag.OrFlags];
            AllFlags = [.. flag.AllFlags];
            IsMultiGroup = flag.IsMultiGroup;
            EnableStage = flag.EnableStage;
        }

        public Flag Copy()
        {
            return new Flag(this);
        }

        private string GetFlagWithoutF()
        {
            return _regex.Replace(FlagName, "");
        }

        public string GetBinTableName()
        {
            return "Bin_" + GetFlagWithoutF() + "_Over_Failing";
        }

        public string GetSetFailCoreCounterName()
        {
            return "Set_" + GetFlagWithoutF() + "_Harvest_FailingCoreCount";
        }

        public string GetFlagVariable()
        {
            return _regex.Replace(FlagName, "");
        }
    }

    public class HeaderInfo
    {
        public int Start = -1;
        public int End = -1;
    }
    public class Fusing
    {
        public string Field { set; get; } = "";
        public string Address { set; get; } = "";
        public string Value { set; get; } = "";

        public Fusing() { }

        public Fusing(Fusing fusing)
        {
            if (fusing == null)
            {
                return;
            }

            Field = fusing.Field;
            Address = fusing.Address;
            Value = fusing.Value;
        }

        public Fusing Copy()
        {
            return new Fusing(this);
        }

        public string GetAddress()
        {
            return Address.Split('=').First().Trim().Trim('\t').Trim();
        }

        public string GetAddressFlag()
        {
            if (!Address.Contains("=>"))
            {
                return "";
            }

            return Address.Split('>').Last().Trim().Trim('\t').Trim();
        }

        public List<string> GetAddressFlags()
        {
            if (!Address.Contains("=>"))
            {
                return [];
            }

            var flags = new List<string>();
            string flag = Address.Split('>').Last().Trim().Trim('\t').Trim();
            GroupItem groupItem = flag.GetIterates();

            if (groupItem.Groups.Count != 0)
            {
                foreach (KeyValuePair<int, List<Tuple<string, string>>> item in groupItem.Groups)
                {
                    string text = flag;
                    foreach (Tuple<string, string> list in item.Value)
                    {
                        text = text.Replace(list.Item1, list.Item2);
                    }

                    flags.Add(text);
                }
                return flags;
            }
            flags.Add(flag);
            return flags;
        }
    }

    public partial class HarvestingTruthTableRow : MyRow
    {
        [GeneratedRegex(@"\[(?<value>\w+)+\]", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"[^\w]", RegexOptions.Compiled)]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"\{(?<value>\w+)+\}", RegexOptions.Compiled)]
        private static partial Regex MyRegex2();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();

        public string Condition { set; get; } = "";
        public string ProposedBinName { set; get; } = string.Empty;
        public string HardBin { set; get; } = "";
        public List<Flag> Flags = [];
        public List<Flag> RepairFlags = [];
        public Dictionary<string, string> OutputBinTables = [];
        public Dictionary<string, string> OutputHarvestDecisionFlags = [];
        public Dictionary<string, string> OutputFlags = [];

        public List<Fusing> Fusings = [];
        public List<Fusing> ReadFuse = [];
        public Dictionary<string, string> FlagSyntaxDic = [];
        #region Constructor
        public HarvestingTruthTableRow(string sheetName = "")
        {
            SheetName = sheetName;
        }

        public HarvestingTruthTableRow(HarvestingTruthTableRow harvestingTruthTableRow) : base(harvestingTruthTableRow)
        {
            if (harvestingTruthTableRow == null)
            {
                return;
            }

            Condition = harvestingTruthTableRow.Condition;
            ProposedBinName = harvestingTruthTableRow.ProposedBinName;
            HardBin = harvestingTruthTableRow.HardBin;
            Flags = [.. harvestingTruthTableRow.Flags.Select(x => x.Copy())];
            RepairFlags = [.. harvestingTruthTableRow.RepairFlags.Select(x => x.Copy())];
            OutputBinTables = new Dictionary<string, string>(harvestingTruthTableRow.OutputBinTables);
            OutputHarvestDecisionFlags = new Dictionary<string, string>(harvestingTruthTableRow.OutputHarvestDecisionFlags);
            OutputFlags = new Dictionary<string, string>(harvestingTruthTableRow.OutputFlags);
            Fusings = [.. harvestingTruthTableRow.Fusings.Select(x => x.Copy())];
            ReadFuse = [.. harvestingTruthTableRow.ReadFuse.Select(x => x.Copy())];
            FlagSyntaxDic = new Dictionary<string, string>(harvestingTruthTableRow.FlagSyntaxDic);
        }

        public HarvestingTruthTableRow Copy()
        {
            return new HarvestingTruthTableRow(this);
        }
        #endregion

        public string GetBinResult()
        {
            string binResult = GetDecisionResult();
            return HardBin ?? binResult;
        }

        public string GetDecisionResult()
        {
            string binResult = "";
            if (OutputBinTables.ContainsValue("1"))
            {
                binResult = OutputBinTables.ToList().Find(x => x.Value == "1").Key;
            }

            return binResult;
        }

        public List<string> GetRepairDomain()
        {
            List<string> binNameItem = [.. ProposedBinName.Split('_')];
            var domainList = new List<string>();
            int harvIdx = binNameItem.FindLastIndex(x => x.StartsWithIgnoreCase("Harv"));
            if (harvIdx == -1)
            {
                return domainList;
            }

            string harvItem = binNameItem[harvIdx];
            if (harvItem.Contains('S'))
            {
                domainList.Add("S");
            }

            if (harvItem.Contains('C'))
            {
                domainList.Add("C");
            }

            if (harvItem.Contains('G'))
            {
                domainList.Add("G");
            }

            return domainList;
        }

        public Dictionary<string, Dictionary<string, string>>? GetProposedBinNameList()
        {
            var binNameList = new Dictionary<string, Dictionary<string, string>>();
            List<string> binNameItem = [.. ProposedBinName.Split('_')];
            int repairIdx = binNameItem.FindLastIndex(x => x.StartsWithIgnoreCase("repair"));
            if (repairIdx != -1)
            {
                binNameItem[repairIdx] = "repair";
            }

            int elementCount = string.Join("_", binNameItem).Count(x => x.Equals(']'));
            var combinList = new List<string>();
            double totalCombination = Math.Pow(2, elementCount);
            for (int i = 0; i < totalCombination; i++)
            {
                combinList.Add(Convert.ToString(i, 2).PadLeft(elementCount, '0'));
            }
            foreach (string combin in combinList)
            {
                var itemList = new List<string>();
                int idx = 0;
                var flagDic = new Dictionary<string, string>();
                var expandSyntax = new Dictionary<string, Dictionary<string, string>>();
                foreach (string item in binNameItem)
                {
                    if (_regex.IsMatch(item))
                    {
                        var replaceList = new Dictionary<string, string>();
                        MatchCollection syntax = _regex.Matches(item);
                        string tmpItem = item;
                        foreach (Match keyword in syntax)
                        {
                            string key = keyword.Groups[1].Value;
                            Flag? targetFlag = Flags.Find(x => x.Syntax == key);
                            if (targetFlag != null)
                            {
                                string tarFlag = targetFlag.FlagName + "_0";
                                string status = combin[idx] == '0' ? "T" : "F";
                                flagDic.Add(tarFlag, status);
                                replaceList.Add(key, combin[idx].ToString());
                                //itemList.Add(item.Replace(key, combin[idx].ToString()).Replace("[", "").Replace("]", ""));
                            }
                            idx++;
                        }
                        foreach (KeyValuePair<string, string> replaceItem in replaceList)
                        {
                            tmpItem = tmpItem.Replace(replaceItem.Key, replaceItem.Value);
                        }

                        itemList.Add(tmpItem.Replace("[", "").Replace("]", ""));
                    }
                    else if (_regex3.IsMatch(item))
                    {
                        MatchCollection syntax = _regex3.Matches(item);
                        string tmpItem = item;
                        foreach (Match keyword in syntax)
                        {
                            string key = keyword.Groups[1].Value;
                            Flag? targetFlag = Flags.Find(x => x.Syntax == key);
                            if (targetFlag != null)
                            {
                                string withoutCore = targetFlag.FlagWithoutCore;
                                var flagMapping = targetFlag.AllFlags.ToDictionary(x => x.Replace(withoutCore, ""), x => x);
                                expandSyntax.Add("{" + key + "}", flagMapping);
                            }
                        }
                        itemList.Add(item);
                    }
                    else
                    {
                        itemList.Add(item);
                    }
                }
                string binName = string.Join("_", itemList);
                binNameList.TryAdd(binName, flagDic);
                ExpandBinName(ref binNameList, expandSyntax);
            }
            return binNameList;
        }

        private static void ExpandBinName(ref Dictionary<string, Dictionary<string, string>> oldBinNameList, Dictionary<string, Dictionary<string, string>> expandSyntax)
        {
            var newBinNameList = new Dictionary<string, Dictionary<string, string>>();
            bool hasMod = false;
            foreach (KeyValuePair<string, Dictionary<string, string>> binName in oldBinNameList)
            {
                string name = binName.Key;
                var syntaxs = expandSyntax.Keys.ToList();
                string? needToReplaceItem = syntaxs.Find(name.Contains);
                if (needToReplaceItem != null)
                {
                    foreach (KeyValuePair<string, string> flag in expandSyntax[needToReplaceItem])
                    {
                        string coreNum = flag.Key;
                        string flagName = flag.Value;
                        string newBinName = name.Replace(needToReplaceItem, coreNum);
                        var newFlagDic = new Dictionary<string, string>();
                        foreach (KeyValuePair<string, string> oldFlag in binName.Value)
                        {
                            newFlagDic.Add(oldFlag.Key, oldFlag.Value);
                        }
                        newFlagDic.Add(flagName, "T");
                        newBinNameList.Add(newBinName, newFlagDic);
                    }
                    hasMod = true;
                }
                else
                {
                    newBinNameList.Add(binName.Key, binName.Value);
                }
            }

            if (hasMod)
            {
                ExpandBinName(ref newBinNameList, expandSyntax);
            }

            oldBinNameList = newBinNameList;
        }

        public string GetModifiedCondition()
        {
            return _regex2.Replace(Condition, "_").Replace(" ", "_");
        }

        public string GetModifiedProposedBinName()
        {
            return ProposedBinName[..ProposedBinName.IndexOf("_repair", StringComparison.OrdinalIgnoreCase)];
        }

        public string GetConditionInstacne(string? jobName = null)
        {
            string condition = GetModifiedCondition() + "_Harvest_eFuse_Write";
            if (!string.IsNullOrEmpty(jobName))
            {
                condition += "_" + jobName;
            }

            return condition;
        }

        public List<string> GetSumFlags()
        {
            var sumFlags = new List<string>();
            foreach (Flag flag in Flags)
            {
                if (flag.IsSum)
                {
                    sumFlags.Add(flag.FlagName);
                    sumFlags.AddRange(flag.SumFlags);
                }
            }
            return sumFlags;
        }

        public string GetHarvGlobalFailFlag()
        {
            var flags = Flags.Where(x => x.IsSum).Select(x => x.AllFlags).ToList();
            return string.Join("|", flags.Select(x => string.Join("+", x)));
        }

        public string GetCustHarvGlobalFailFlag()
        {
            var flags = Flags.Where(x => x.IsCust).Select(x => x.Name).ToList();
            return string.Join("|", flags);
        }

        public string HarvAllCorePassFlag()
        {
            var flags = Flags.Where(x => x.IsSum).Select(x => x.FlagName + "_0").ToList();
            return string.Join("|", flags);
        }

        public string HarvBinOutFailFlag()
        {
            var flags = Flags.Where(x => x.IsSum).Select(x => x.FlagName).ToList();
            return string.Join("|", flags);
        }

        public string CustHarvBinOutFailFlag()
        {
            var flags = Flags.Where(x => x.IsCust).Select(x => x.FlagName).ToList();
            return string.Join("|", flags);
        }
    }
}
