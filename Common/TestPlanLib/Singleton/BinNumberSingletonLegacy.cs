using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using OfficeOpenXml;

using TestPlanLib.BinNumberLegacy;

namespace TestPlanLib.Singleton
{
    public sealed class BinNumberSingletonLegacy
    {
        #region Singleton
        private static BinNumberSingletonLegacy? _instance;
        private static string _binNumberConfig = "";
        private static string _currentProject = "";
        private static readonly HashSet<string> _binTableNames = [];

        private List<SoftBinDigiDef> SoftBinDigiDef { get; set; } = [];
        private List<SoftBinRangeData> SoftBinRangeDatas { get; set; } = [];
        private ExcelWorkbook BinNumerWorkbook { get; }
        private List<int> UsebinList { get; } = [9995, 9200, 9300, 9400, 9500, 9100, 9800, 9998];
        private int Last { get; set; }

        private BinNumberSingletonLegacy()
        {
            if (File.Exists(_binNumberConfig))
            {
                string binNumberPath = _binNumberConfig;
                BinNumerWorkbook = new ExcelPackage(new FileInfo(binNumberPath)).Workbook;
            }
            else
            {
                string defaultBinNumber = Path.Combine(AppContext.BaseDirectory, "Config", "BinNumberConfig.xlsx");
                string binNumberPath = Path.Combine(AppContext.BaseDirectory, "Config", $"BinNumberConfig_{_currentProject}.xlsx");
                if (File.Exists(defaultBinNumber))
                {
                    File.Copy(defaultBinNumber, binNumberPath, true);
                }
                BinNumerWorkbook = new ExcelPackage(new FileInfo(binNumberPath)).Workbook;
            }

            if (BinNumerWorkbook == null || BinNumerWorkbook.Worksheets.Count == 0)
            {
                throw new Exception("Can not find Bin number config file, please check if it has existed!");
            }

            InitialSoftBinDigit();

            InitialSoftBinRange();
        }

        public static BinNumberSingletonLegacy Instance()
        {
            return _instance ??= new BinNumberSingletonLegacy();
        }
        #endregion

        #region Initialization
        private void InitialSoftBinDigit()
        {
            ExcelWorksheet softBinWorksheet = BinNumerWorkbook.Worksheets["Bin_Number_Def"];
            ExcelWorksheet moduleSheet = BinNumerWorkbook.Worksheets["ModuleList"];
            var reader = new SoftBinDataReader(BinNumerWorkbook);
            reader.ReadModuleList(moduleSheet);
            SoftBinDigiDef = reader.ReadSheet(softBinWorksheet);
        }

        private void InitialSoftBinRange()
        {
            ExcelWorksheet rangeWorksheet = BinNumerWorkbook.Worksheets["DC_HardIp_Bin_Def"];
            var reader = new SoftBinRangeReader();
            SoftBinRangeDatas = reader.ReadSheet(rangeWorksheet);
        }
        #endregion

        public static void Initialize(string binNumberConfig, string currentProject)
        {
            _instance = null;
            _binNumberConfig = binNumberConfig;
            _currentProject = currentProject;
        }

        public void SetStartBinNumber(int newStartBinNum)
        {
            Last = newStartBinNum;
        }

        public string GetEfuseSoftBinNumber(BinTableRow binTableRow)
        {
            int binNum = Last + 1;
            while (UsebinList.Contains(binNum))
            {
                binNum++;
            }

            Last = binNum > 149 ? 149 : binNum;
            return binNum.ToString();
        }

        public string GetSoftBinNumber(Dictionary<string, string> modes, Action<string, EnumMessageLevel, int, string> responseReport, bool alarm, string binName, SoftBinNumPara softBinNumPara, bool isBinCut = false)
        {
            string softBin = "9999";
            foreach (SoftBinDigiDef binDigiDef in SoftBinDigiDef)
            {
                bool match = false;
                switch (binDigiDef.CategoryType)
                {
                    case EnumBinNumKeyType.Category:
                        match = softBinNumPara.Category.EqualsIgnoreCase(binDigiDef.Category);
                        break;

                    case EnumBinNumKeyType.Module:
                        match = softBinNumPara.Module.EqualsIgnoreCase(binDigiDef.Category);
                        break;

                    case EnumBinNumKeyType.SubModule:
                        match = softBinNumPara.SubModule.EqualsIgnoreCase(binDigiDef.Category);
                        break;

                    case EnumBinNumKeyType.Block:
                        match = softBinNumPara.Block.EqualsIgnoreCase(binDigiDef.Category);
                        break;
                }
                if (match)
                {
                    if (isBinCut)
                    {
                        binDigiDef.GetBinNumberBinCut(softBinNumPara, out softBin);
                    }
                    else
                    {
                        binDigiDef.GetBinNumber(modes, softBinNumPara, out softBin);
                    }
                }
            }
            if (softBin == "9999" && !_binTableNames.Contains(binName.ToUpper()))
            {
                string errorMessage;
                if (softBinNumPara.ColumnContentDic?.ContainsKey("Label") == true
                    && softBinNumPara.ColumnContentDic.TryGetValue("Pattern", out string? value) && softBinNumPara.ColumnContentDic.TryGetValue("Voltage", out string? value1)
                     && softBinNumPara.ColumnContentDic.TryGetValue("Level", out string? value2))
                {
                    errorMessage =
                        $"The softBin of {binName} is 9999, Label : {softBinNumPara.ColumnContentDic["Label"]} , Pattern : {value} , Voltage : {value1}  , Level : {value2}";
                }
                else
                {
                    errorMessage = $"The softBin of {binName} is 9999";
                }

                if (alarm)
                {
                    responseReport(errorMessage, EnumMessageLevel.Error, 10, "");
                }
            }
            _binTableNames.Add(binName.ToUpper());
            int binValue = Convert.ToInt32(softBin);
            if (!UsebinList.Exists(x => x == binValue))
            {
                UsebinList.Add(binValue);
            }

            return softBin;
        }

        public bool GetBinNumDefRow(BinNumDefPara binNumDefPara, out BinNumDefRow binNumDefRow)
        {
            binNumDefRow = new BinNumDefRow();
            SoftBinRangeData? targetItem = SearchSoftBinRangeData(binNumDefPara);
            if (targetItem != null)
            {
                binNumDefRow.Description = targetItem.Description;
                binNumDefRow.CurrentSoftBin = targetItem.GetSoftBinNumber();
                binNumDefRow.HardBin = targetItem.HardBin;
                binNumDefRow.SoftBinStart = targetItem.GetSoftBinStart();
                binNumDefRow.SoftBinEnd = targetItem.GetSoftBinEnd();
                binNumDefRow.SoftBinState = targetItem.GetStatus();
                binNumDefRow.IsExceed = targetItem.CheckExceed();
                binNumDefRow.CurrentBinLib = targetItem;
                binNumDefRow.HardIphlvBin = targetItem.HardHlvBin;
                binNumDefRow.HardIpHvBin = targetItem.HardHvBin;
                binNumDefRow.HardIpLvBin = targetItem.HardLvBin;
                binNumDefRow.HardIpNvBin = targetItem.HardNvBin;
                return true;
            }
            return false;
        }

        public bool GetHardBinNum(BinNumDefPara binNumDefPara, out string hardBinNum)
        {
            hardBinNum = "";
            var binNumForBlock = SoftBinRangeDatas.Where(p => p.Block.Replace(" ", "").EqualsIgnoreCase(binNumDefPara.Block.Replace(" ", "")))
                .Select(a => a).ToList();

            if (binNumForBlock.Exists(p => p.Match(binNumDefPara.Condition)))
            {
                SoftBinRangeData targetItem = binNumForBlock.Find(p => p.Match(binNumDefPara.Condition))!;
                hardBinNum = targetItem.HardBin;
                return true;
            }
            return false;
        }

        public SoftBinRangeData? SearchSoftBinRangeData(BinNumDefPara binNumDefPara)
        {
            var binNumForBlock = SoftBinRangeDatas
                .Where(p => p.Block.EqualsIgnoreCase(binNumDefPara.Block))
                .Select(a => a).ToList();

            SoftBinRangeData? result = null;
            if (binNumForBlock.Exists(p => p.Match(binNumDefPara.Condition)))
            {
                result = binNumForBlock.Find(p => p.Match(binNumDefPara.Condition));
            }

            if (result == null)
            {
                binNumForBlock = [.. SoftBinRangeDatas.Where(p => p.Block.EqualsIgnoreCase(nameof(EnumBinNumDefBlock.Default))).Select(a => a)];
                result = binNumForBlock.Find(p => p.Match("Default"));
            }
            return result;
        }
    }
}
