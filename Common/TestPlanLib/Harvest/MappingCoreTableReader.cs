using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace TestPlanLib.Harvest
{
    public class MappingCoreTableReader : MySheetReader<MappingCoreTable>
    {
        private const string ConHeaderPattern = "Pattern";
        private const string ConHeaderHarvestPinGrp = "HarvestPinGrp";
        private const string ConHeaderHarvestPinFlag = "HarvestPinFlag";

        private int _indexPattern = -1;
        private int _indexHarvestPinGrp = -1;
        private int _indexHarvestPinFlag = -1;
        private string _sheetName = "";
        private MappingCoreTable? _sheet;
        public override MappingCoreTable? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;
            _sheet = new MappingCoreTable(_sheetName);
            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet(_sheetName);
            // _error.Add(ErrorReportManager.ErrorInstance);

            return _sheet!;
        }

        private void ReadSheet(string sheetName)
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new MappingCoreRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexPattern != -1)
                {
                    row.Pattern = ExcelWorksheet.GetCellValue(i, _indexPattern).Trim();
                }

                if (row.Pattern.Length == 0)
                {
                    continue;
                }

                if (_indexHarvestPinGrp != -1 && _indexHarvestPinFlag != -1)
                {
                    row.PinGrpFlagDic = GetPinGrpFlags(ExcelWorksheet.GetCellValue(i, _indexHarvestPinGrp).Trim(), ExcelWorksheet.GetCellValue(i, _indexHarvestPinFlag).Trim(), i);
                    row.IsFullPattern =
                        string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(i, _indexHarvestPinGrp).Trim());
                }
                if (_sheet!.Rows.Exists(x => x.Pattern.EqualsIgnoreCase(row.Pattern)))
                {
                    string errorMessage = $"Duplicate Pattern : {row.Pattern}";
                    _sheet!.AddError(SsnErrorType.E_DuplicatePattern_01, _sheetName, i, _indexPattern, $"Duplicate Pattern : {row.Pattern}", [row.Pattern]);
                }
                _sheet!.Rows.Add(row);
            }

            _sheet!.IndexPattern = _indexPattern;
            _sheet!.IndexHarvestPinGrp = _indexHarvestPinGrp;
            _sheet!.IndexHarvestPinFlag = _indexHarvestPinFlag;
        }

        private Dictionary<string, string> GetPinGrpFlags(string pinGrps, string pinFlags, int rowNum)
        {
            var pinGrpFlagDic = new Dictionary<string, string>();
            string[] pinGrpList = pinGrps.Split(';');
            string[] pinFlagList = pinFlags.Split(';');
            if (pinGrpList.Length != pinFlagList.Length)
            {
                //string errorMessage = $"PinGroup Count : {pinGrpList.Length} , PinFlag Count : {pinFlagList.Length} Mismatch";
                _sheet!.AddError(SsnErrorType.E_MismatchFlag_01, _sheetName, rowNum, _indexHarvestPinFlag, $"PinGroup Count : {pinGrpList.Length} , PinFlag Count : {pinFlagList.Length} Mismatch", [pinGrpList.Length.ToString(), pinFlagList.Length.ToString()]);
            }
            if (pinGrpList.Length != 0)
            {
                for (int i = 0; i < pinGrpList.Length; i++)
                {
                    if (pinFlagList.Length <= i)
                    {
                        continue;
                    }

                    string flagName = pinFlagList[i];
                    string pinGroupName = pinGrpList[i];
                    if (!flagName.StartsWithIgnoreCase("F_"))
                    {
                        //string errorMessage = $"Illegal Flag Name : {flagName}";
                        _sheet!.AddError(SsnErrorType.E_InvalidFlag_01, _sheetName, rowNum, _indexHarvestPinFlag, $"Illegal Flag Name : {flagName}", [flagName]);
                    }
                    if (!pinGrpFlagDic.TryAdd(pinGroupName, flagName))
                    {
                        //string errorMessage = $"Duplicate Pin Group : {pinGroupName}";
                        _sheet!.AddError(SsnErrorType.E_DuplicatePinGroup_01, _sheetName, rowNum, _indexHarvestPinGrp, $"Duplicate Pin Group : {pinGroupName}", [pinGroupName]);
                    }
                }
            }
            else if (string.IsNullOrEmpty(pinGrps) && !string.IsNullOrEmpty(pinFlags))
            {
                pinGrpFlagDic.Add("", pinFlags);
            }
            return pinGrpFlagDic;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderPattern))
                {
                    _indexPattern = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderHarvestPinGrp))
                {
                    _indexHarvestPinGrp = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderHarvestPinFlag))
                {
                    _indexHarvestPinFlag = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderPattern))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class MappingCoreTable : MySheet
    {
        public List<Regex> RegexPlList;
        public List<MappingCoreRow> Rows { set; get; }

        public int IndexPattern = -1;
        public int IndexHarvestPinGrp = -1;
        public int IndexHarvestPinFlag = -1;
        #region Constructor
        public MappingCoreTable(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
            RegexPlList = [];
        }
        #endregion
        public void CreateRegexList()
        {
            foreach (MappingCoreRow row in Rows.Where(x => x.IsSsn()))
            {
                var regexPattern = new Regex($"{row.Pattern.ToLower().Replace("*", "(.+)?")}", RegexOptions.Compiled);
                RegexPlList.Add(regexPattern);
            }
        }

        public static SubFlowSheet CreateCoreMaskSubFlow()
        {
            var subflow = new SubFlowSheet("Flow_SSN_CoreMask");
            subflow.AddRow(new FlowRow { Opcode = OpCode.Test, Parameter = "SSN_CoreMask" });
            subflow.AddReturnRow();
            return subflow;
        }
    }

    public partial class MappingCoreRow : MyRow
    {
        [GeneratedRegex("(SSC|SSU)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public string Pattern = "";
        public Dictionary<string, string> PinGrpFlagDic = [];
        public bool IsFullPattern;
        public string InitPattern = "";
        public string CoreName = "";
        public string HarvestFlag = "";
        public string PowerSupply = "";
        public string Comment = "";
        #region Constructor
        public MappingCoreRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

        public bool IsSsn()
        {
            return _regex.IsMatch(Pattern);
        }
    }
}
