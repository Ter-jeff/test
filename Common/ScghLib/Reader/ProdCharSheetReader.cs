using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using OfficeOpenXml;

using ScghLib.Base;

namespace ScghLib.Reader
{
    public partial class ProdCharSheetReader : MySheetReader<ProdCharSheet>
    {
        protected const string ConHeaderBlock = "Block";
        protected const string ConHeaderMode = "MODE";
        protected const string ConHeaderItem = "ITEM";
        protected const string ConHeaderApplication = "APPLICATION";
        protected const string ConHeaderPayload = @"PAYLOAD\d*";
        protected const string ConHeaderSupplyVoltage = "SUPPLY VOLTAGE";
        protected const string ConHeaderInit = @"INIT\d+";
        protected const string ConHeaderUsage = "USAGE.*";
        protected const string ConHeaderComment = "COMMENT.*";
        protected const string ConPeripheralVoltage = "Peripheral Voltage";
        protected const string ConSramVoltage = "SRAM Voltage";
        protected const string ConHeaderEnable = "Enable";
        protected const string ConHeaderLevelHVorLv = "Level.*";

        [GeneratedRegex(ConHeaderPayload, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(ConHeaderInit, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"\n")]
        private static partial Regex MyRegex2();

        protected int BlockColNumber;
        protected int ModeColNumber;
        protected int ItemColNumber;
        protected int ApplicationColNumber = -1;
        protected int InitStartColNumber;
        protected int PayloadStartColNumber;
        protected int UsageColNumber = -1;
        protected int SupplyVoltageColNumber;
        protected int CommentColNumber;
        protected int SVoltageColNumber;
        protected int PVoltageColNumber;

        protected string Module = string.Empty;
        protected int InitCnt;
        protected int PayloadCnt;
        protected string SheetName = string.Empty;
        protected int EnableWordColumnNumber = -1;
        protected int LevelHVorLv = -1;

        public override ProdCharSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;
            SheetName = excelWorksheet.Name;

            ReadHeader();

            ProdCharSheet sheetScgScan = ReadAllData();

            sheetScgScan.SheetName = excelWorksheet.Name;

            return sheetScgScan;
        }

        /// <summary>
        /// This function is to get the number of Init in the sheet, the Init header
        /// can be Init1, Init2 ..., we will find all the header between "Application" and "Payload", 
        /// and count the number of all the header, return it.
        /// </summary>
        protected virtual void ReadHeader()
        {
            ResetValue();

            GetDimensions();

            GetFirstHeaderPosition();

            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).ToUpper().Trim();

                if (CellDiff.IsLiked(header, ConHeaderBlock))
                {
                    BlockColNumber = i;
                    continue;
                }
                if (CellDiff.IsLiked(header, ConHeaderMode))
                {
                    ModeColNumber = i;
                    continue;
                }
                if (CellDiff.IsLiked(header, ConHeaderItem))
                {
                    ItemColNumber = i;
                    continue;
                }
                if (CellDiff.IsLiked(header, ConHeaderApplication))
                {
                    ApplicationColNumber = i;
                    continue;
                }

                if (CellDiff.IsLiked(header, ConHeaderSupplyVoltage))
                {
                    SupplyVoltageColNumber = i;
                    continue;
                }

                if (MyRegex().IsMatch(header))
                {
                    if (PayloadStartColNumber == 0)
                    {
                        PayloadStartColNumber = i;
                    }
                    PayloadCnt++;
                    continue;
                }

                if (MyRegex1().IsMatch(header))
                {
                    if (InitStartColNumber == 0)
                    {
                        InitStartColNumber = i;
                    }
                    InitCnt++;
                    continue;
                }

                if (CellDiff.IsLiked(header, ConHeaderEnable))
                {
                    EnableWordColumnNumber = i;
                    continue;
                }

                if (CellDiff.IsLiked(header, ConHeaderLevelHVorLv))
                {
                    LevelHVorLv = i;
                    continue;
                }

                if (CellDiff.IsLiked(header, ConHeaderUsage))
                {
                    UsageColNumber = i;
                    continue;
                }
                if (CellDiff.IsLiked(header, ConHeaderComment))
                {
                    CommentColNumber = i;
                    continue;
                }
                if (CellDiff.IsLiked(header, ConPeripheralVoltage))
                {
                    PVoltageColNumber = i;
                    continue;
                }
                if (CellDiff.IsLiked(header, ConSramVoltage))
                {
                    SVoltageColNumber = i;
                }
            }
        }

        /// <summary>
        /// ResetValue
        /// </summary>
        protected void ResetValue()
        {
            StartCol = 1;
            StartRow = 1;
            EndCol = 1;
            EndRow = 1;
            InitCnt = 0;
            InitStartColNumber = 0;

            EnableWordColumnNumber = -1;
            LevelHVorLv = -1;

            UsageColNumber = -1;
            ApplicationColNumber = -1;
        }

        /// <summary>
        /// Get the position of First header which defined in the field
        /// </summary>
        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderMode))
                    {
                        StartRow = i;
                        break;
                    }

                }
            }
        }

        /// <summary>
        /// read the init value for one row in the input sheet
        /// </summary>
        /// <param name="rowNumber">rowNumber</param>
        /// <returns></returns>
        protected List<string> ReadInitList(int rowNumber)
        {
            var initList = new List<string>();
            for (int i = 0; i < InitCnt; i++)
            {
                initList.Add(ExcelWorksheet.GetCellValue(rowNumber, InitStartColNumber + i).Trim());
            }
            return initList;
        }

        protected List<string> ReadPayloadList(int rowNumber)
        {
            var payloadList = new List<string>();
            for (int i = 0; i < PayloadCnt; i++)
            {
                string payload = ExcelWorksheet.GetCellValue(rowNumber, PayloadStartColNumber + i).Trim();
                if (!string.IsNullOrEmpty(payload) && payload != "NA")
                {
                    payloadList.Add(payload);
                }
            }
            return payloadList;
        }

        protected virtual ProdCharSheet ReadAllData()
        {
            var sheetScg = new ProdCharSheet();
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var scgRow = new ProdCharSheetRow(SheetName);
                if (!string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(i, PayloadStartColNumber)))
                {
                    scgRow.RowNum = i;
                    scgRow.Block = BlockColNumber == 0 ? "" : ExcelWorksheet.GetCellValue(i, BlockColNumber).Trim();
                    scgRow.Mode = ExcelWorksheet.GetCellValue(i, ModeColNumber).Trim();
                    scgRow.Item = ExcelWorksheet.GetCellValue(i, ItemColNumber).Trim();

                    if (PayloadStartColNumber != 0)
                    {
                        List<string> payloads = ReadPayloadList(i);
                        scgRow.PayloadList.AddRange(payloads);
                        scgRow.PayloadAliasList.AddRange(payloads);
                    }

                    if (InitStartColNumber != 0)
                    {
                        List<string> inits = ReadInitList(i);
                        scgRow.InitList.AddRange(inits);
                        scgRow.InitAliasList.AddRange(inits);
                    }

                    scgRow.Application = ApplicationColNumber != -1 ? ExcelWorksheet.GetCellValue(i, ApplicationColNumber).Trim() : GetApplicationField(scgRow.PayloadValue);

                    scgRow.Usage = UsageColNumber != -1 ? ExcelWorksheet.GetCellValue(i, UsageColNumber).Trim() : "1";

                    if (SupplyVoltageColNumber != 0)
                    {
                        scgRow.SupplyVoltage = ExcelWorksheet.GetCellValue(i, SupplyVoltageColNumber).Trim();
                        scgRow.SupplyVoltage = MyRegex2().Replace(scgRow.SupplyVoltage, "");
                    }

                    if (SVoltageColNumber != 0 && PVoltageColNumber != 0)
                    {
                        scgRow.SramVoltage = ExcelWorksheet.GetCellValue(i, SVoltageColNumber).Trim();
                        scgRow.SramVoltage = MyRegex2().Replace(scgRow.SramVoltage, "");
                        scgRow.SramVoltage = scgRow.SramVoltage.EqualsIgnoreCase("N/A") ? "" : scgRow.SramVoltage;
                        scgRow.SramVoltage = scgRow.SramVoltage.EqualsIgnoreCase("NA") ? "" : scgRow.SramVoltage;

                        scgRow.PeripheralVoltage = ExcelWorksheet.GetCellValue(i, PVoltageColNumber).Trim();
                        scgRow.PeripheralVoltage = MyRegex2().Replace(scgRow.PeripheralVoltage, "");
                        scgRow.PeripheralVoltage = scgRow.PeripheralVoltage.EqualsIgnoreCase("N/A") ? "" : scgRow.PeripheralVoltage;
                        scgRow.PeripheralVoltage = scgRow.PeripheralVoltage.EqualsIgnoreCase("NA") ? "" : scgRow.PeripheralVoltage;
                    }

                    if (EnableWordColumnNumber != -1)
                    {
                        scgRow.EnableWord = ExcelWorksheet.GetCellValue(i, EnableWordColumnNumber).Trim();
                    }

                    if (LevelHVorLv != -1)
                    {
                        scgRow.LevelHVorLv = ExcelWorksheet.GetCellValue(i, LevelHVorLv).Trim();
                    }

                    sheetScg.RowList.Add(scgRow);
                }
            }
            return sheetScg;
        }

        protected static string GetApplicationField(string payLoadName)
        {
            string appName = "";
            if (string.IsNullOrEmpty(payLoadName))
            {
                return appName;
            }

            string[] payloadTok = payLoadName.ToUpper().Split(['_'], StringSplitOptions.RemoveEmptyEntries);
            if (payloadTok.Length != 0)
            {
                if (payloadTok[0] == "PP")
                {
                    appName = "Production";
                }
                else if (payloadTok[0] == "CZ")
                {
                    appName = "Characterization";
                }
                else if (payloadTok[0] == "DD")
                {
                    appName = "Debug";
                }
                else if (payloadTok[0] == "HT")
                {
                    appName = "HTOL";
                }
                else
                {
                    appName = payloadTok[0];
                }
            }
            return appName;
        }

    }

    public partial class ProdCharSheet
    {
        private static readonly Regex _regex1 = MyRegex();
        private static readonly Regex _regex = MyRegex1();

        [GeneratedRegex(@"_DSRM\d|_SRM\d", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("_DSRMDSSC|_SRMDSSC", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        public string SheetName { get; set; } = string.Empty;
        public List<ProdCharSheetRow> RowList { get; set; } = [];

        protected static List<string> FindInitInMode(string init, List<ProdCharSheetRow> prodCharSheetRows)
        {
            var rows = prodCharSheetRows.FindAll(p => p.Mode.EqualsIgnoreCase(init)).Select(p => p.PayloadValue).ToList();
            return rows;
        }

        protected static List<string> FindInitInItem(string init, List<ProdCharSheetRow> prodCharSheetRows)
        {
            return [.. prodCharSheetRows.FindAll(p => p.Item.EqualsIgnoreCase(init)).Select(p => p.PayloadValue)];
        }

        public List<List<string>> GetPatternLists()
        {
            var initLists = new List<List<string>>();
            foreach (ProdCharSheetRow row in RowList)
            {
                var initList = new List<string>();
                for (int index = 0; index < row.InitList.Count; index++)
                {
                    string init = row.InitList[index];
                    if (!string.IsNullOrEmpty(init) && init != "NA" && init != "N/A")
                    {
                        List<string> inits = FindInitInMode(init, RowList);
                        if (inits.Count == 0)
                        {
                            inits = FindInitInItem(init, RowList);
                        }

                        if (inits.Count == 0)
                        {
                            inits.Add(init);
                        }

                        initList.AddRange(inits);
                    }
                    else
                    {
                        initList.Add("");
                    }
                }
                initList.AddRange(row.PayloadList);
                initLists.Add(initList);
            }
            return initLists;
        }

        public Dictionary<string, List<string>> GetSelSramInits()
        {
            var dic = new Dictionary<string, List<string>>();
            var dsscInits = new Dictionary<string, List<string>>();
            var hcInits = new Dictionary<string, List<string>>();
            foreach (ProdCharSheetRow row in RowList)
            {
                var initLists = new List<List<string>>();
                for (int index = 0; index < row.InitList.Count; index++)
                {
                    string init = row.InitList[index];
                    var initList = new List<string>();
                    if (!string.IsNullOrEmpty(init) && init != "NA" && init != "N/A")
                    {
                        List<string> inits = FindInitInMode(init, RowList);
                        if (inits.Count == 0)
                        {
                            inits = FindInitInItem(init, RowList);
                        }

                        if (inits.Count == 0)
                        {
                            inits.Add(init);
                        }

                        initList.AddRange(inits);
                    }
                    else
                    {
                        initList.Add("");
                    }
                    initLists.Add(initList);
                }

                List<List<string>> combos = Combination.GetAllPossibleCombos(initLists);
                foreach (List<string> inits in combos)
                {
                    if (dsscInits.Count == 0)
                    {
                        Dictionary<string, List<string>> dssc = MatchPattern(inits, "_DSSC", _regex);
                        if (dssc.Count != 0)
                        {
                            dsscInits = dssc;
                        }
                    }
                    if (hcInits.Count == 0)
                    {
                        Dictionary<string, List<string>> hc = MatchPattern(inits, "_HC", _regex1);
                        if (hc.Count != 0)
                        {
                            hcInits = hc;
                        }
                    }
                    if (dsscInits.Count != 0 && hcInits.Count != 0)
                    {
                        foreach (KeyValuePair<string, List<string>> init in dsscInits)
                        {
                            dic.Add(init.Key, init.Value);
                        }

                        foreach (KeyValuePair<string, List<string>> init in hcInits)
                        {
                            dic.Add(init.Key, init.Value);
                        }

                        return dic;
                    }
                }
            }

            if (dsscInits.Count != 0)
            {
                foreach (KeyValuePair<string, List<string>> init in dsscInits)
                {
                    dic.Add(init.Key, init.Value);
                }

                return dic;
            }
            if (hcInits.Count != 0)
            {
                foreach (KeyValuePair<string, List<string>> init in hcInits)
                {
                    dic.Add(init.Key, init.Value);
                }
            }

            return dic;
        }

        private Dictionary<string, List<string>> MatchPattern(List<string> inits, string text, Regex regex)
        {
            var dicInitPatterns = new Dictionary<string, List<string>>();
            int index = inits.FindIndex(regex.IsMatch);
            if (index != -1)
            {
                string key = SheetName + text;
                if (!dicInitPatterns.ContainsKey(key))
                {
                    dicInitPatterns.Add(key, [.. inits.GetRange(0, index + 1).Where(x => !string.IsNullOrEmpty(x))]);
                }
            }
            return dicInitPatterns;
        }
    }

    public partial class ProdCharSheetRow : IProdCharItem
    {
        private const string ConCpu = "Cpu";
        private const string ConGpu = "Gfx";
        private const string ConSoc = "Soc";
        private const string ConSpi = "Spi";

        public int RowNum { set; get; }
        public string Block { set; get; } = string.Empty;
        public string Mode { set; get; } = string.Empty;
        public string Item { set; get; } = string.Empty;
        public string Segment { set; get; } = string.Empty;
        public string Inits { set; get; } = string.Empty;
        public string PayLoads { set; get; } = string.Empty;
        public bool IsGenFlow { set; get; }
        public string Application { set; get; } = string.Empty;
        public List<string> PatternList { set; get; } = [];
        public List<string> InitList { set; get; } = [];
        public List<string> PayloadList { set; get; } = [];
        public List<string> InitAliasList { set; get; } = [];
        public List<string> PayloadAliasList { set; get; } = [];
        public string Usage { set; get; } = string.Empty;
        public string SupplyVoltage { set; get; } = string.Empty;
        public string EnableWord { set; get; } = string.Empty;
        public string LevelHVorLv { set; get; } = string.Empty;
        public string Comments { set; get; } = string.Empty;
        public string PeripheralVoltage { get; set; } = string.Empty;
        public string SramVoltage { get; set; } = string.Empty;
        public string SourceSheetName { get; } = string.Empty;
        public string FlowName { set; get; } = string.Empty;
        public string ForceCondition { get; set; } = string.Empty;
        public string DcUsed { get; set; } = string.Empty;
        public string EnableHlnv { get; set; } = string.Empty;

        public string PayloadValue
        {
            get
            {
                if (PayloadList.Count == 0)
                {
                    return "";
                }

                return PayloadList.First();
            }
        }

        public string Chiplet
        {
            get
            {
                return ProdCharSheetRowHelpers.MyRegex().Match(SourceSheetName).Groups["chiplet"].ToString();
            }
        }

        public ProdCharSheetRow(string sourceSheetName = "")
        {
            SourceSheetName = sourceSheetName;
            Block = "";
            Mode = "";
            Item = "";
            Segment = "";
            Inits = "";
            PayLoads = "";
            IsGenFlow = true;
            Application = "";
            PatternList = [];
            InitList = [];
            PayloadList = [];
            InitAliasList = [];
            PayloadAliasList = [];
            Usage = "";
            SupplyVoltage = "";
            EnableWord = "";
            LevelHVorLv = "";
            Comments = "";
            PeripheralVoltage = "";
            SramVoltage = "";
            ForceCondition = "";
            DcUsed = "";
            EnableHlnv = "";
        }

        public ProdCharSheetRow(ProdCharSheetRow prodCharSheetRow)
        {
            if (prodCharSheetRow == null)
            {
                return;
            }

            SourceSheetName = prodCharSheetRow.SourceSheetName;
            RowNum = prodCharSheetRow.RowNum;
            Block = prodCharSheetRow.Block;
            Mode = prodCharSheetRow.Mode;
            Item = prodCharSheetRow.Item;
            Segment = prodCharSheetRow.Segment;
            Inits = prodCharSheetRow.Inits;
            PayLoads = prodCharSheetRow.PayLoads;
            IsGenFlow = prodCharSheetRow.IsGenFlow;
            Application = prodCharSheetRow.Application;
            Usage = prodCharSheetRow.Usage;
            SupplyVoltage = prodCharSheetRow.SupplyVoltage;
            EnableWord = prodCharSheetRow.EnableWord;
            LevelHVorLv = prodCharSheetRow.LevelHVorLv;
            Comments = prodCharSheetRow.Comments;
            PeripheralVoltage = prodCharSheetRow.PeripheralVoltage;
            SramVoltage = prodCharSheetRow.SramVoltage;
            FlowName = prodCharSheetRow.FlowName;
            ForceCondition = prodCharSheetRow.ForceCondition;
            DcUsed = prodCharSheetRow.DcUsed;
            EnableHlnv = prodCharSheetRow.EnableHlnv;

            PatternList = prodCharSheetRow.PatternList?.ToList() ?? [];
            InitList = prodCharSheetRow.InitList?.ToList() ?? [];
            PayloadList = prodCharSheetRow.PayloadList?.ToList() ?? [];
            InitAliasList = prodCharSheetRow.InitAliasList?.ToList() ?? [];
            PayloadAliasList = prodCharSheetRow.PayloadAliasList?.ToList() ?? [];
        }

        public ProdCharSheetRow Copy()
        {
            return new ProdCharSheetRow(this);
        }

        public List<string> GetInitList()
        {
            return [.. InitList.Where(p => !string.IsNullOrEmpty(p))];
        }

        public List<string> GetPayloadList()
        {
            return [.. PayloadList.Where(p => !string.IsNullOrEmpty(p))];
        }

        public List<string> GetInitAliasList()
        {
            return [.. InitAliasList.Where(p => !string.IsNullOrEmpty(p))];
        }

        public List<string> GetPayloadAliasList()
        {
            return [.. PayloadAliasList.Where(p => !string.IsNullOrEmpty(p))];
        }

        public string GetIndexForHardIp()
        {
            List<string> validInitList = InitList.FindAll(s => !string.IsNullOrEmpty(s) && s != "NA" && s != "N/A");
            List<string> validPayLoadList = PayloadList.FindAll(s => !string.IsNullOrEmpty(s) && s != "NA" && s != "N/A");
            if (InitList.Count > 1 || PayloadList.Count > 1)
            {
                return (string.Join(",", validInitList) + ";" + string.Join(";", validPayLoadList)).Trim(';');
            }
            return PayloadValue;
        }

        public string GetSourceSheet()
        {
            return SourceSheetName;
        }

        public List<string> GetAllPatternList()
        {
            List<string> all = [.. InitList, .. PayloadList];
            return all;
        }

        public string GetDomainByPattern()
        {
            string domain = "";
            //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'
            //private static readonly Regex RgxOrg = new Regex(@"A|C|L|P|S|V|H", RegexOptions.IgnoreCase | RegexOptions.Compiled); //DP for Dummy Pattern

            foreach (string pattern in PatternList)
            {
                List<string> arr = [.. pattern.Split('_')];
                if (arr.Count > 2)
                {
                    if (arr[2].EqualsIgnoreCase("A"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("P"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("V"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("H"))
                    {
                        domain = "HardIP";
                    }
                    else if (arr[2].EqualsIgnoreCase("C"))
                    {
                        domain = "Cpu";
                    }
                    else if (arr[2].EqualsIgnoreCase("L"))
                    {
                        domain = "Gfx";
                    }
                    else if (arr[2].EqualsIgnoreCase("S"))
                    {
                        domain = "Soc";
                    }
                }
                if (!string.IsNullOrEmpty(domain))
                {
                    return domain;
                }
            }
            return "";
        }

        public string GetDomainByFlowName()
        {
            string domain = "";
            if (FlowName.ContainsIgnoreCase(ConCpu.ToLower()))
            {
                domain = ConCpu;
            }
            else if (FlowName.ContainsIgnoreCase(ConGpu.ToLower()) ||
                FlowName.ContainsIgnoreCase("gpu") ||
                FlowName.ContainsIgnoreCase("gfx"))
            {
                domain = ConGpu;
            }
            else if (FlowName.ContainsIgnoreCase(ConSoc.ToLower()))
            {
                domain = ConSoc;
            }
            else if (FlowName.ContainsIgnoreCase(ConSpi.ToLower()))
            {
                domain = ConSpi;
            }

            return domain;
        }
    }
}
