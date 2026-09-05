using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

using OfficeOpenXml;

using TestPlanLib.BinNumber;

namespace TestPlanLib.Singleton
{
    public sealed class BinNumberSingleton
    {
        private const string BinNumber = "BinNumber";
        private static BinNumberSingleton? _instance;
        public static BinNumberSingleton Instance
        {
            get
            {
                return _instance ??= new BinNumberSingleton();
            }
        }
        private static Dictionary<string, BinNumResult> _bincutBinNums = [];
        public static Dictionary<string, BinNumResult> BincutBinNums => _bincutBinNums ??= [];
        private Dictionary<int, BinNumUsedStatus> _allBinNumUsedStatus = [];
        private HashSet<BinNumInfo> _binNumInfoSet = [];
        private SortedSet<int> _freeSoftBinNums = [];
        private Dictionary<string, BinNumResult> _existBinRows = [];

        private BinNumberSingleton()
        {
            Initialize();
        }

        public void Initialize(ExcelWorkbook? excelWorkbook = null)
        {
            _allBinNumUsedStatus = [];
            _binNumInfoSet = [];
            _freeSoftBinNums = [];
            _existBinRows = [];
            _bincutBinNums = [];

            var allDefinedSoftBinNums = new SortedSet<int>();
            for (int i = 1; i <= 9999; i++)
            {
                _allBinNumUsedStatus[i] = new BinNumUsedStatus();
            }
            if (excelWorkbook != null)
            {
                foreach (ExcelWorksheet sheet in excelWorkbook.Worksheets)
                {
                    if (sheet.Name.EqualsIgnoreCase(BinNumber))
                    {
                        _binNumInfoSet = new BinNumSheetReader().ReadSheet(sheet);
                    }
                }
            }
            if (_binNumInfoSet.Count != 0)
            {
                allDefinedSoftBinNums = [.. _binNumInfoSet.SelectMany(x => x.SoftBinNums)];
            }
            _freeSoftBinNums = [.. _allBinNumUsedStatus.Keys.Except(allDefinedSoftBinNums)];
        }

        public bool CheckDuplicateBinRows(BinTableRow binTableRow)
        {
            if (binTableRow != null && _existBinRows.ContainsKey(binTableRow.GetUniqKey()))
            {
                return true;
            }
            return false;
        }

        public BinNumResult GetBinInfo(string module, string category1, string category2, BinTableRow? binTableRow = null, bool queryOnly = false)
        {
            if (binTableRow != null && _existBinRows.TryGetValue(binTableRow.GetUniqKey(), out BinNumResult? cachedResult))
            {
                return cachedResult;
            }

            string cleanModule = module?.Trim() ?? string.Empty;
            string cleanCat1 = NormalizeCategory(category1);
            string cleanCat2 = NormalizeCategory(category2);

            KeyValuePair<int, BinNumUsedStatus> binNumUsedStatus;
            BinNumInfo binNumInfo;
            bool found = false;
            if (_binNumInfoSet.Count != 0)
            {
                (binNumUsedStatus, binNumInfo, found) = SearchModuleBinNum(cleanModule, cleanCat1, cleanCat2);
            }
            else
            {
                (binNumUsedStatus, binNumInfo) = SearchDefaultBinNum(cleanModule, cleanCat1, cleanCat2);
            }

            if (!queryOnly)
            {
                if (!_allBinNumUsedStatus.ContainsKey(binNumUsedStatus.Key))
                {
                    _allBinNumUsedStatus.Add(binNumUsedStatus.Key, binNumUsedStatus.Value);
                }
                binNumUsedStatus.Value.Used = true;
                binNumUsedStatus.Value.UsedBinNumInfos.Add(binNumInfo);
            }
            if (binTableRow != null)
            {
                _existBinRows.Add(binTableRow.GetUniqKey(), new BinNumResult(binNumUsedStatus.Key, binNumInfo, found));
            }

            return new BinNumResult(binNumUsedStatus.Key, binNumInfo, found);
        }

        private static string NormalizeCategory(string cat)
        {
            if (string.IsNullOrWhiteSpace(cat))
            {
                return string.Empty;
            }

            string trimmed = cat.Trim();

            if (trimmed.EqualsIgnoreCase("X"))
            {
                return string.Empty;
            }

            return trimmed.Contains('_') ? trimmed.Replace("_", "") : trimmed;
        }

        private (KeyValuePair<int, BinNumUsedStatus>, BinNumInfo) SearchDefaultBinNum(string module, string category1, string category2)
        {
            KeyValuePair<int, BinNumUsedStatus> target = _allBinNumUsedStatus.Where(x => !x.Value.Used).OrderBy(x => x.Key).FirstOrDefault();
            if (target.Value == null)
            {
                target = _allBinNumUsedStatus.OrderBy(x => x.Key).LastOrDefault();
            }
            return (target, CreateBinNumSearchKey(module, category1, category2));
        }

        private (KeyValuePair<int, BinNumUsedStatus>, BinNumInfo, bool) SearchModuleBinNum(string module, string category1, string category2)
        {
            bool found = false;
            BinNumInfo searchKey1 = CreateBinNumSearchKey(module, category1, category2);
            BinNumInfo searchKey2 = CreateBinNumSearchKey(module, category1, "");
            BinNumInfo searchKey3 = CreateBinNumSearchKey(module, "", "");

            if (_binNumInfoSet.TryGetValue(searchKey1, out BinNumInfo? binNumInfo))
            {
                found = true;
            }
            else if (_binNumInfoSet.TryGetValue(searchKey2, out binNumInfo))
            {
                found = true;
            }
            else if (_binNumInfoSet.TryGetValue(searchKey3, out binNumInfo))
            {
                found = true;
            }
            else
            {
                binNumInfo = null;
            }

            KeyValuePair<int, BinNumUsedStatus> target;
            SortedSet<int> ints;
            if (binNumInfo != null)
            {
                ints = binNumInfo.SoftBinNums;
            }
            else
            {
                ints = _freeSoftBinNums;
                binNumInfo = CreateBinNumSearchKey(module, category1, category2);
            }

            var free = _allBinNumUsedStatus.Where(x => !x.Value.Used).Select(x => x.Key).Intersect(ints).OrderBy(x => x).ToList();

            if (free.Count != 0)
            {
                target = new KeyValuePair<int, BinNumUsedStatus>(free.FirstOrDefault(), _allBinNumUsedStatus[free.FirstOrDefault()]);
            }
            else
            {
                target = _allBinNumUsedStatus.ContainsKey(ints.LastOrDefault()) ? new KeyValuePair<int, BinNumUsedStatus>(ints.LastOrDefault(), _allBinNumUsedStatus[ints.LastOrDefault()]) : new KeyValuePair<int, BinNumUsedStatus>(ints.LastOrDefault(), new BinNumUsedStatus());
            }

            return (target, binNumInfo, found);
        }

        private static BinNumInfo CreateBinNumSearchKey(string module, string category1, string category2)
        {
            return new BinNumInfo()
            {
                Module = module,
                Category1 = category1,
                Category2 = category2
            };
        }
    }
}
