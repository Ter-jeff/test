using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace Automation.Utility.TpUpdate.HardIPEnableWordsUpdate
{
    public class EnableWordTable
    {
        private const string PattHeader = "Pattern";
        private const string CatHeader = "Category";
        private const string SubblockHeader = "SubBlock";
        private const string FailControlHeader = "HIP_FailRun";

        private static readonly List<TtrColumnConfig> _headerConfig =
        [
            new(PattHeader, true),
            new(CatHeader, true),
            new(SubblockHeader, true),
            new(FailControlHeader, false)
        ];

        public bool IsContainsSubBlock()
        {
            if (_headerMetadata == null)
            {
                return false;
            }
            return _headerMetadata.Any(m => m.Name == SubblockHeader);
        }

        internal IEnumerable<TtrColumnMetadata> _headerMetadata;

        /// <summary>
        /// TTR table is only considered valid when "Proposal for TTR" cell exist 
        /// and there is vaild header after "Proposal for TTR" row
        /// </summary>
        public bool IsValid { get; private set; }

        public List<string> EnableWords { get; set; } = [];

        public Dictionary<string, List<string>> OverlayInfoDic { get; private set; } = [];

        public List<EnableWordTableRow> Rows { get; set; } = [];

        /// <summary>
        /// Load TTR table from Excel worksheet. <see cref="IsValid"/> will be set to true 
        /// if "Proposal for TTR" cell exist and there is vaild header after "Proposal for TTR" row
        /// </summary>
        /// <param name="sheet"></param>
        public void ReadTable(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null)
            {
                return;
            }
            int startRow = GetValidStartingRowIndex(sheet);
            if (startRow == -1)
            {
                return;
            }
            int headerRowIndex = GetValidHeaderRowIndex(sheet, startRow);
            if (headerRowIndex == -1)
            {
                return;
            }
            IsValid = true;

            IEnumerable<TtrColumnMetadata> headerMetadata = GetHeaderMetadata(sheet, headerRowIndex);
            _headerMetadata = headerMetadata;

            List<EnableWordTableRow> rows = GetBodyData(sheet, headerRowIndex + 1, headerMetadata);
            Rows = rows;
            OverlayInfoDic = GetBincutOverlay(rows);
            EnableWords = rows.SelectMany(p => p.TotalEnableWords).Distinct().ToList();
            var enablesTrim = new List<string>();
            foreach (string enable in EnableWords)
            {
                enablesTrim.Add(enable.Replace(" ", "_"));
            }
            EnableWords = GenJobAllEnableWord(enablesTrim);
        }

        internal static List<EnableWordTableRow> GetBodyData(ExcelWorksheet sheet, int startRow, IEnumerable<TtrColumnMetadata> headerMetadata)
        {
            Dictionary<string, TtrColumnMetadata> metadataDictionary = headerMetadata
                .ToDictionary(m => m.Name, m => m);

            List<EnableWordTableRow> rows = [];

            // Load table rows from excel worksheet
            for (int i = startRow; i <= sheet.Dimension.End.Row; i++)
            {
                var datas = new Dictionary<int, string>();
                foreach (TtrColumnMetadata column in headerMetadata)
                {
                    datas.Add(column.Index, EpplusExtensions.GetCellValue(sheet, i, column.Index));
                }
                if (datas.All(p => p.Value == ""))
                {
                    break;
                }

                int patternIndex = headerMetadata.Single(m => m.Name == PattHeader).Index;
                int subBlockIndex = headerMetadata.Single(m => m.Name == SubblockHeader).Index;

                string mainkey = datas[patternIndex];
                string subKey = datas[subBlockIndex];
                if (string.IsNullOrEmpty(subKey))
                {
                    continue;
                }

                EnableWordTableRow itemkey = rows.FirstOrDefault(p =>
                    p.Name.Equals(mainkey, StringComparison.OrdinalIgnoreCase)
                    && p.SubBlock.Equals(subKey, StringComparison.OrdinalIgnoreCase)
                );
                if (itemkey != null)
                {
                    rows.Remove(itemkey);
                }
                var row = new EnableWordTableRow(PattHeader, CatHeader, SubblockHeader, FailControlHeader, metadataDictionary, datas)
                {
                    RowNum = i + 1
                };
                rows.Add(row);
            }

            return rows.OrderByDescending(p => p.Name).ToList();
        }

        internal static Dictionary<string, List<string>> GetBincutOverlay(List<EnableWordTableRow> rows)
        {
            Dictionary<string, List<string>> overlayInfoDic = [];
            List<EnableWordTableRow> needBincutOverlayRows = rows
                .FindAll(x => x.EnableWords.ContainsKey("LV_PBcut") || x.EnableWords.ContainsKey("HV_PBCut"));
            if (needBincutOverlayRows.Count == 0)
            {
                return overlayInfoDic;
            }
            foreach (EnableWordTableRow row in needBincutOverlayRows)
            {
                if (row.Name.Split('_').Length <= 9)
                {
                    continue;
                }
                string mode = row.Name.Split('_')[9].ToUpper().Replace("X", "");

                if (!overlayInfoDic.ContainsKey(mode))
                {
                    overlayInfoDic.Add(mode, []);
                }

                if (row.EnableWords.TryGetValue("LV_PBCut", out List<string> lvPbCut))
                {
                    overlayInfoDic[mode].AddRange(lvPbCut);
                }
                if (row.EnableWords.TryGetValue("HV_PBCut", out List<string> hvPbCut))
                {
                    overlayInfoDic[mode].AddRange(hvPbCut);
                }
                overlayInfoDic[mode] = overlayInfoDic[mode].Distinct().ToList();
            }
            return overlayInfoDic;
        }

        internal static IEnumerable<TtrColumnMetadata> GetHeaderMetadata(ExcelWorksheet sheet, int rowIndex)
        {
            Dictionary<string, int> names = EpplusExtensions.GetHeaderOrder(sheet, rowIndex);
            return names
                .Where(name => !Regex.IsMatch(name.Key, "HV|LV|NV", RegexOptions.IgnoreCase))
                .Select(pair => new TtrColumnMetadata(pair.Key, pair.Value));
        }

        private static int GetValidHeaderRowIndex(ExcelWorksheet sheet, int startRow)
        {
            for (int i = startRow; i <= sheet.Dimension.End.Row; i++)
            {
                if (IsRowValidHeader(sheet, i))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Check is target row contains required header columns. 
        /// All column definition is in <see cref="_headerConfig"/>
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">Target row index</param>
        /// <returns></returns>
        private static bool IsRowValidHeader(ExcelWorksheet sheet, int rowIndex)
        {
            Dictionary<string, int> names = EpplusExtensions.GetHeaderOrder(sheet, rowIndex);
            foreach (TtrColumnConfig colConf in _headerConfig.Where(cf => cf.Required))
            {
                bool exist = names.ContainsKey(colConf.Name);
                if (!exist)
                {
                    return false;
                }
            }
            return true;
        }

        private const int FirstRowIndex = 1;

        /// <summary>
        /// Find row index of "Proposal for TTR"
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        internal static int GetValidStartingRowIndex(ExcelWorksheet sheet)
        {
            for (int i = FirstRowIndex; i <= sheet.Dimension.End.Row; i++)
            {
                Dictionary<string, int> names = EpplusExtensions.GetHeaderOrder(sheet, i);
                foreach (string str in names.Keys)
                {
                    if (Regex.IsMatch(str, "Proposal for TTR", RegexOptions.IgnoreCase))
                    {
                        return i + 1;
                    }
                }
            }
            return -1;
        }

        internal static List<string> GenJobAllEnableWord(List<string> allEnableWords)
        {
            string regstr = @"(?<Enable>[a-zA-Z]+)_*\s*(?<Job>\w*(cp)|(ft)|(wlft))";
            var testEnables = new List<string>();
            List<string> itemsTotal = allEnableWords.FindAll(p => Regex.IsMatch(p, regstr, RegexOptions.IgnoreCase));
            //Use Job to Search all enable words
            var enableList = new List<string>();
            foreach (string item in itemsTotal)
            {
                string enableitem = Regex.Match(item, regstr, RegexOptions.IgnoreCase).Groups["Enable"].ToString();
                if (!enableList.Contains(enableitem))
                {
                    enableList.Add(enableitem);
                }
            }
            //Use Enable Words to search all jobs with specified Enable Word
            foreach (string enable in enableList)
            {
                testEnables.Add(enable + "_All");
            }
            allEnableWords.AddRange(testEnables);
            return allEnableWords;
        }
    }

    public class EnableWordTableRow : MyRow
    {
        /// <summary>
        /// Pattern name
        /// </summary>
        public string Name { get; set; }

        public string Category { get; set; }

        public string SubBlock { get; set; }

        public string FailControl { get; set; }

        public bool IsOther
        {
            get { return SubBlock.Equals("other", StringComparison.OrdinalIgnoreCase); }
        }

        public List<string> TotalEnableWords { get; private set; } = new List<string>();

        public Dictionary<string, List<string>> EnableWords { get; set; } =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        public EnableWordTableRow(string patternName, string categoryName, string subblock, string failControlName, Dictionary<string, TtrColumnMetadata> headerMetadata, Dictionary<int, string> dataset)
        {
            int categoryIndex = headerMetadata[categoryName].Index;
            int subBlockIndex = headerMetadata[subblock].Index;
            int patternIndex = headerMetadata[patternName].Index;
            Category = dataset[categoryIndex];
            SubBlock = dataset[subBlockIndex];
            Name = dataset[patternIndex];

            bool failRunExist = headerMetadata
                .TryGetValue(failControlName, out TtrColumnMetadata failRunMetadata);
            if (failRunExist)
            {
                FailControl = dataset[failRunMetadata.Index];
            }
            else
            {
                FailControl = "";
            }

            ConvertDatasetToEnableWords(headerMetadata, dataset);
        }

        private void ConvertDatasetToEnableWords(Dictionary<string, TtrColumnMetadata> headerMetadata, Dictionary<int, string> dataset)
        {
            foreach (KeyValuePair<string, TtrColumnMetadata> index in headerMetadata)
            {
                string items = dataset[index.Value.Index];
                if (items == Name || items == Category)
                {
                    continue;
                }

                string indexKey = index.Key.Replace(" ", "_");
                TotalEnableWords.Add(indexKey);
                if (items == "" && !IsValidHlnSyntax(items))
                {
                    continue;
                }

                items = items.Replace(" ", "");
                if (items == "")
                {
                    continue;
                }
                foreach (string item in items.Split(','))
                {
                    string key = item.Trim();
                    if (!EnableWords.ContainsKey(key))
                    {
                        EnableWords.Add(key, [indexKey]);
                    }
                    else
                    {
                        EnableWords[key].Add(indexKey);
                    }
                }
            }
        }

        internal bool IsValidHlnSyntax(string str)
        {
            bool result = true;
            string regHln = "^[HLN]V$";
            string regBv = "^H?BV$";
            string regPbCut = "^[L|H]V_PBCut";
            string regChk = "^H?BV_Chk";
            List<string> items = str.Replace(" ", "").Split(',').ToList();
            foreach (string checkItem in items)
            {
                if (!Regex.IsMatch(checkItem, regHln, RegexOptions.IgnoreCase) && !Regex.IsMatch(checkItem, regBv, RegexOptions.IgnoreCase) && !Regex.IsMatch(checkItem, regPbCut, RegexOptions.IgnoreCase) && !Regex.IsMatch(checkItem, regChk, RegexOptions.IgnoreCase))
                {
                    result = false;
                    break;
                }
            }
            return result;
        }
    }

    public record TtrColumnConfig(string Name, bool Required);

    public record TtrColumnMetadata(string Name, int Index);
}
