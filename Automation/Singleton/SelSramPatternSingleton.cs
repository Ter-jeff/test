using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.SelSram;
using Automation.Static;

using CommonLib.Extension;

using OfficeOpenXml;

using ScghLib.Reader;

using TestPlanLib.BinCut;
using TestPlanLib.Static;

namespace Automation.Singleton
{
    public class SelSramPatternSingleton
    {
        private static SelSramPatternSingleton _instance;

        private const string ConHc = "HC";
        private const string ConDssc = "DSSC";

        public Dictionary<string, List<string>> DicInitPatterns { get; set; }
        public List<SelSramData> SelSramDatas;
        private static readonly Regex _regex = new Regex("_DSRM[0-9a-fA-F]+$|_SRM[0-9a-fA-F]+$|_DSELSRM[0-9a-fA-F]+$|_SELSRM[0-9a-fA-F]+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex("_DSRMDSSC$|_SRMDSSC$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public HashSet<string> Chiplet { get; set; }
        public HashSet<string> BitsLogicPins { get; set; }
        public List<string> ReadBackBits { get; set; }
        public Dictionary<string, string> DicReadbackPat { get; set; }
        public Dictionary<string, List<string>> DicReadbackInitPat { get; set; }
        public bool HasReadBackInitPat { get; set; }
        public Dictionary<string, OtherPatData> DicOtherPat { get; set; }
        public List<string> CombosWrite { get; set; }
        public List<string> CombosRead { get; set; }
        public List<string> CombosCheck { get; set; }

        private SelSramPatternSingleton()
        {
            Chiplet = new HashSet<string> { "" };
            SelSramDatas = new List<SelSramData>();
            CombosWrite = new List<string>();
            CombosRead = new List<string>();
            CombosCheck = new List<string>();
            BitsLogicPins = new HashSet<string>();
            ReadBackBits = new List<string>();
            DicReadbackPat = new Dictionary<string, string>();
            DicReadbackInitPat = new Dictionary<string, List<string>>();
            DicOtherPat = new Dictionary<string, OtherPatData>();
            DicInitPatterns = new Dictionary<string, List<string>>();
        }

        #region Singleton
        public static SelSramPatternSingleton GetInstance()
        {
            return _instance ?? (_instance = new SelSramPatternSingleton());
        }
        #endregion

        public void FillData()
        {
            SelSramDatas.Clear();

            var prodCharSheets = new List<ProdCharSheet>();
            List<string> sheets = NeededSheets.GetProdCharSheets();
            var scghSheets = sheets.Where(x => x.ContainsIgnoreCase("_SC") ||
                x.ContainsIgnoreCase("_BI") ||
                x.ContainsIgnoreCase("SCAN") ||
                x.ContainsIgnoreCase("BIST")
                ).ToList();
            if (EpWorkbook.ScghWorkbook == null)
            {
                return;
            }

            foreach (string scghSheet in scghSheets)
            {
                ExcelWorksheet worksheet = EpWorkbook.ScghWorkbook.Worksheets[scghSheet];
                if (worksheet == null)
                {
                    continue;
                }

                var reader = new ProdCharSheetReader();
                ProdCharSheet prodCharSheet = reader.ReadSheet(worksheet);
                prodCharSheets.Add(prodCharSheet);
            }

            foreach (ProdCharSheet prodCharSheet in prodCharSheets)
            {
                string pcSheetName = prodCharSheet.SheetName;
                foreach (ProdCharSheetRow row in prodCharSheet.RowList)
                {
                    foreach (string pattern in row.InitList)
                    {
                        if (pattern.Split('_').Length < 8)
                        {
                            continue;
                        }

                        // HardCode
                        if (_regex.IsMatch(pattern))
                        {
                            AddData(new SelSramData(pcSheetName, pattern, ConHc));
                        }
                        // DSSC
                        else if (_regex2.IsMatch(pattern))
                        {
                            AddData(new SelSramData(pcSheetName, pattern, ConDssc));
                        }
                    }

                    foreach (string pattern in row.PayloadList)
                    {
                        if (pattern.Split('_').Length < 8)
                        {
                            continue;
                        }

                        // HardCode
                        if (_regex.IsMatch(pattern))
                        {
                            AddData(new SelSramData(pcSheetName, pattern, ConHc));
                        }
                        // DSSC
                        else if (_regex2.IsMatch(pattern))
                        {
                            AddData(new SelSramData(pcSheetName, pattern, ConDssc));
                        }
                    }
                }
                Dictionary<string, List<string>> dic = HasReadBackInitPat ? GetInitDicFromReadback(prodCharSheet) : prodCharSheet.GetSelSramInits();
                if (dic != null)
                {
                    foreach (KeyValuePair<string, List<string>> item in dic)
                    {
                        if (!DicInitPatterns.ContainsKey(item.Key))
                        {
                            DicInitPatterns.Add(item.Key, item.Value);
                        }
                    }
                }
            }
        }

        private (string, string) GetOrgByPcSheetName(string pcSheetName)
        {
            var dic = new Dictionary<string, List<string>>
            {
                { "SOCSCAN", NeededSheets.ScanScghSoc.Split(',').ToList() },
                { "SOCMBIST", NeededSheets.MbistCharScgSoc.Split(',').ToList() },
                { "CPUSCAN", NeededSheets.ScanScghCpu.Split(',').ToList() },
                { "CPUMBIST", NeededSheets.MbistCharScgCpu.Split(',').ToList() },
                { "GFXSCAN", NeededSheets.ScanScghGpu.Split(',').ToList() },
                { "GFXMBIST", NeededSheets.MbistCharScgGpu.Split(',').ToList() }
            };

            KeyValuePair<string, List<string>> target = dic.FirstOrDefault(x => x.Value.Any(y => y.Equals(pcSheetName, StringComparison.OrdinalIgnoreCase)));
            if (!string.IsNullOrEmpty(target.Key))
            {
                if (Regex.IsMatch(pcSheetName, @"_[a-zA-z]\d$"))//chiplet
                {
                    return ($"{target.Key}_{pcSheetName.Split('_').LastOrDefault()}", pcSheetName.Split('_').LastOrDefault());
                }

                return (target.Key, "");
            }
            return ("", "");
        }

        private Dictionary<string, List<string>> GetInitDicFromReadback(ProdCharSheet prodCharSheet)
        {
            var dic = new Dictionary<string, List<string>>();
            (string targetOrgName, string chiplet) = GetOrgByPcSheetName(prodCharSheet.SheetName);
            KeyValuePair<string, List<string>> targetOrg = DicReadbackInitPat.FirstOrDefault(x => x.Key.Equals(targetOrgName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(targetOrg.Key))
            {
                Chiplet.Add(chiplet.Trim().ToUpper());
                SelSramData targetHcPattern = SelSramDatas.FirstOrDefault(x => x.Category.Equals(prodCharSheet.SheetName, StringComparison.OrdinalIgnoreCase) && x.Type == ConHc);
                SelSramData targetDsscPattern = SelSramDatas.FirstOrDefault(x => x.Category.Equals(prodCharSheet.SheetName, StringComparison.OrdinalIgnoreCase) && x.Type == ConDssc);

                if (targetHcPattern != null)
                {
                    List<string> initPats = new List<string>(targetOrg.Value)
                    {
                        targetHcPattern.Pattern
                    };
                    dic.Add(targetHcPattern.Category + "_HC", initPats);
                }
                if (targetDsscPattern != null)
                {
                    List<string> initPats = new List<string>(targetOrg.Value)
                    {
                        targetDsscPattern.Pattern
                    };
                    dic.Add(targetDsscPattern.Category + "_DSSC", initPats);
                }
            }
            return dic;
        }

        public void Write2Txt(string output)
        {
            var selsramMappingTable = new SelsrmMappingSheet();
            ExcelWorksheet sheet = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRM_Mapping_Table"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Mapping_Table"];
            if (sheet != null)
            {
                var selsrmMappingSheetReader = new SelsrmMappingSheetReader();
                selsramMappingTable = selsrmMappingSheetReader.ReadSheet(sheet);
            }
            IOrderedEnumerable<string> socScanSheets = NeededSheets.ScanScghSoc.Split(',').OrderBy(x => x);
            IOrderedEnumerable<string> socBistSheets = NeededSheets.MbistCharScgSoc.Split(',').OrderBy(x => x);
            IOrderedEnumerable<string> cpuScanSheets = NeededSheets.ScanScghCpu.Split(',').OrderBy(x => x);
            IOrderedEnumerable<string> cpuBistSheets = NeededSheets.MbistCharScgCpu.Split(',').OrderBy(x => x);
            IOrderedEnumerable<string> gpuScanSheets = NeededSheets.ScanScghGpu.Split(',').OrderBy(x => x);
            IOrderedEnumerable<string> gpuBistSheets = NeededSheets.MbistCharScgGpu.Split(',').OrderBy(x => x);

            var allSheets = new List<string>();
            allSheets.AddRange(socScanSheets);
            allSheets.AddRange(socBistSheets);
            allSheets.AddRange(cpuScanSheets);
            allSheets.AddRange(cpuBistSheets);
            allSheets.AddRange(gpuScanSheets);
            allSheets.AddRange(gpuBistSheets);
            var allDataColDict = new Dictionary<string, List<string>>();

            foreach (string pcSheet in allSheets)
            {
                List<SelSramData> data = SelSramDatas.FindAll(p => p.Category.Equals(pcSheet, StringComparison.OrdinalIgnoreCase) && p.Type.Equals(ConHc));
                if (data.Any())
                {
                    string keyName = $"{pcSheet.ToUpper().Replace("_PC", "")}_{ConHc}";
                    if (!allDataColDict.ContainsKey(keyName))
                    {
                        allDataColDict.Add(keyName, data.Select(x => x.Pattern).ToList());
                    }
                }
                data = SelSramDatas.FindAll(p => p.Category.Equals(pcSheet, StringComparison.OrdinalIgnoreCase) && p.Type.Equals(ConDssc));
                if (data.Any())
                {
                    string keyName = $"{pcSheet.ToUpper().Replace("_PC", "")}_{ConDssc}";
                    if (!allDataColDict.ContainsKey(keyName))
                    {
                        allDataColDict.Add(keyName, data.Select(x => x.Pattern).ToList());
                    }
                }
            }
            allDataColDict.Add("Logic_Pins", BitsLogicPins.ToList());
            allDataColDict.Add("Sram_Pins", new List<string>());
            foreach (string logicPin in BitsLogicPins)
            {
                allDataColDict["Sram_Pins"].Add(selsramMappingTable.Rows.FirstOrDefault(x => x.LogicPins.Equals(logicPin, StringComparison.CurrentCultureIgnoreCase))?.SramPins ?? "");
            }
            allDataColDict.Add("Write", CombosWrite.Select(x => "digsrc" + x).ToList());
            allDataColDict.Add("Read", CombosRead.Select(x => "digcap" + x).ToList());
            allDataColDict.Add("Check", CombosCheck);
            allDataColDict.Add("Src_Cap_bits", new List<string>());

            foreach (var wrPair in CombosWrite.Zip(CombosRead, (w, r) => new { W = w, R = r }))
            {
                allDataColDict["Src_Cap_bits"].Add(GetSrcCapBitsLength(wrPair.W, wrPair.R));
            }
            allDataColDict.Add("Comment", new List<string>());

            var allLines = new List<string> { string.Join("\t", allDataColDict.Keys) };
            int maxRowCount = allDataColDict.Values.Max(x => x.Count);
            for (int i = 0; i < maxRowCount; ++i)
            {
                string line = "";
                foreach (KeyValuePair<string, List<string>> data in allDataColDict)
                {
                    line += (data.Value.ElementAtOrDefault(i) ?? "") + "\t";
                }
                allLines.Add(line);
            }
            File.WriteAllLines(output, allLines);
        }

        private string GetSrcCapBitsLength(string write, string read)
        {
            return write.Length + "," + read.Length;
        }

        public void ReadbackSheet()
        {
            var inputSelSram = new InputSelSram();
            inputSelSram.ParsingReadbackSheet();
        }

        public void AddCombosWrite(string writeStr)
        {
            CombosWrite.Add(writeStr);
        }

        public void AddCombosRead(string readStr)
        {
            CombosRead.Add(readStr);
        }

        public void AddCombosCheck(string checkStr)
        {
            CombosCheck.Add(checkStr);
        }

        public void Add2BitsLogicPins(string pinName)
        {
            BitsLogicPins.Add(pinName);
            ReadBackBits.Add(pinName);
        }

        internal void AddData(SelSramData data)
        {
            if (!SelSramDatas.Exists(p => p.Category.Equals(data.Category) && p.Pattern.Equals(data.Pattern) && p.Type.Equals(data.Type)))
            {
                SelSramDatas.Add(data);
            }
        }

        public static void Initialize()
        {
            _instance = null;
        }
    }
}
