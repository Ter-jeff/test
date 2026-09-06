using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Datalog;
using CommonLib.Extension;

using EfuseCheckCmdLib.DataStructure;
using EfuseCheckCmdLib.Utility;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib.Datalog
{
    public partial class DatalogReader(List<string> lines)
    {
        public static readonly Regex RegEnableWords = EnableWordsRegex();
        public static readonly Regex RegEfuseLine = EfuseLineRegex();
        public static readonly Regex RegPrrLine = PrrLineRegex();
        public static readonly Regex RegSetWriteVariableLine = SetWriteVariableLineRegex();
        public static readonly Regex RegSiteSortBin = SiteSortBinRegex();
        public static readonly Regex RegXY = XYRegex();
        public static readonly Regex RegPRR = PRRRegex();
        public static readonly Regex RegPRRCode = PRRCodeRegex();
        public static readonly Regex RegReadWaferData = ReadWaferDataRegex();
        public static readonly Regex RegLimite = LimiteRegex();
        public static readonly Regex RegIsNumeric = IsNumericRegex();

        [GeneratedRegex(@"Active EnableWords  :", RegexOptions.IgnoreCase)]
        private static partial Regex EnableWordsRegex();

        [GeneratedRegex(@"^Site\(\d+\) EFUSE (Write|Read)(?:\s+.+)?\s+Values", RegexOptions.IgnoreCase)]
        private static partial Regex EfuseLineRegex();

        [GeneratedRegex(@"^\[INFO\]\s+\[Site\s+\d+\] Set PRR", RegexOptions.IgnoreCase)]
        private static partial Regex PrrLineRegex();

        [GeneratedRegex(@"Fuse SetWriteVariable_SiteAware", RegexOptions.IgnoreCase)]
        private static partial Regex SetWriteVariableLineRegex();

        [GeneratedRegex(@"Site\s+Sort\s+Bin", RegexOptions.IgnoreCase)]
        private static partial Regex SiteSortBinRegex();

        [GeneratedRegex(@"Site\s+X_Coord\s+Y_Coord", RegexOptions.IgnoreCase)]
        private static partial Regex XYRegex();

        [GeneratedRegex(@"Set PRR to", RegexOptions.IgnoreCase)]
        private static partial Regex PRRRegex();

        [GeneratedRegex(@"Hex (?:is|are)\s+([0-9A-Fa-f]+)", RegexOptions.IgnoreCase)]
        private static partial Regex PRRCodeRegex();

        [GeneratedRegex(@"<ReadWaferData>", RegexOptions.IgnoreCase)]
        private static partial Regex ReadWaferDataRegex();

        [GeneratedRegex(@"\s+Number\s+Site\s+Test Name\s+Pin", RegexOptions.IgnoreCase)]
        private static partial Regex LimiteRegex();

        [GeneratedRegex(@"^\s*\d+(\s+\d+)*\s*")]
        private static partial Regex IsNumericRegex();

        [GeneratedRegex(@"\s+")]
        private static partial Regex WhitespaceRegex();

        private readonly List<string> _lines = lines;

        public List<DiceInfo> Read(EfuseDramTable efuseDramTable)
        {
            string currentJobStage = "";
            int jobNum = -1;
            string prr = "";

            string scenario = "";
            string dramType = "";
            List<EfuseRow> efuseRows = [];
            List<PrrRow> prrRows = [];
            List<LimitRow> limitRows = [];
            List<SetWriteVariableLine> setWriteVariableLines = [];
            var allDices = new List<DiceInfo>();
            var binInfos = new Dictionary<string, BinInfo>();
            var xyInfos = new Dictionary<string, CoordInfo>();
            ParsingDatalog(efuseDramTable, ref currentJobStage, ref jobNum, ref prr, ref scenario, ref dramType, efuseRows, prrRows, limitRows, setWriteVariableLines, binInfos, xyInfos);
            Console.WriteLine("Parsing datalog Done.");

            SetDiceInfo(currentJobStage, jobNum, scenario, dramType, efuseRows, prrRows, limitRows, setWriteVariableLines, allDices, binInfos, xyInfos);

            return allDices;
        }

        private void ParsingDatalog(EfuseDramTable efuseDramTable, ref string currentJobStage, ref int jobNum, ref string prr, ref string scenario, ref string dramType, List<EfuseRow> efuseRows, List<PrrRow> prrRows, List<LimitRow> limitRows, List<SetWriteVariableLine> setWriteVariableLines, Dictionary<string, BinInfo> binInfos, Dictionary<string, CoordInfo> xyInfos)
        {
            bool foundSiteSortBin = false;
            bool foundXY = false;
            bool hasLimitStart = false;
            bool hasReadWaferData = false;
            int startIndex = 0;
            List<string> lines = _lines;
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                if (TryConsumeLookaheadRow(line, ref foundSiteSortBin, ref foundXY, binInfos, xyInfos))
                {
                    continue;
                }

                if (TryHandleJobNameLine(line, ref currentJobStage, ref jobNum))
                {
                    continue;
                }

                DispatchDataLine(lines, ref i, line, efuseDramTable, ref scenario, ref dramType, ref hasReadWaferData, ref hasLimitStart, ref startIndex, efuseRows, prrRows, limitRows, setWriteVariableLines, ref prr, ref foundSiteSortBin, ref foundXY);
            }
        }

        // The "Site Sort Bin" / "X_Coord Y_Coord" table headers are followed by a run of
        // whitespace-delimited numeric data rows; this consumes those rows while active and
        // reports whether the caller should skip the rest of its per-line handling for this line.
        private static bool TryConsumeLookaheadRow(string line, ref bool foundSiteSortBin, ref bool foundXY, Dictionary<string, BinInfo> binInfos, Dictionary<string, CoordInfo> xyInfos)
        {
            if (!foundSiteSortBin && !foundXY)
            {
                return false;
            }

            string[] dieTest = WhitespaceRegex().Split(line.Trim());
            if (dieTest.Length >= 3 && RegIsNumeric.IsMatch(line))
            {
                if (foundSiteSortBin)
                {
                    var binInfo = new BinInfo { Sort = dieTest[1], Bin = dieTest[2] };
                    binInfos.Add(dieTest[0], binInfo);
                }
                if (foundXY)
                {
                    var xyInfo = new CoordInfo { X = dieTest[1], Y = dieTest[2] };
                    xyInfos.Add(dieTest[0], xyInfo);
                }
                return true;
            }

            if (foundSiteSortBin)
            {
                foundSiteSortBin = false;
            }

            if (foundXY)
            {
                foundXY = false;
            }
            return false;
        }

        private static bool TryHandleJobNameLine(string line, ref string currentJobStage, ref int jobNum)
        {
            if (line.Contains("Job Name:"))
            {
                string jobStage = line.Split(["Job Name:"], StringSplitOptions.None).Last().Trim();
                currentJobStage = jobStage;
                jobNum = EfuseCmdUtility.JobStages.Where(x => jobStage.Contains(x.Key)).Select(x => x.Value).DefaultIfEmpty(999).First();
                return true;
            }

            if (line.Contains("Current job name          :"))
            {
                string jobStage = line.Split(["Current job name", "'"], StringSplitOptions.RemoveEmptyEntries).Last().Trim();
                currentJobStage = jobStage;
                jobNum = EfuseCmdUtility.JobStages.Where(x => jobStage.Contains(x.Key)).Select(x => x.Value).DefaultIfEmpty(999).First();
                return true;
            }

            return false;
        }

        private static void DispatchDataLine(List<string> lines, ref int i, string line, EfuseDramTable efuseDramTable, ref string scenario, ref string dramType, ref bool hasReadWaferData, ref bool hasLimitStart, ref int startIndex, List<EfuseRow> efuseRows, List<PrrRow> prrRows, List<LimitRow> limitRows, List<SetWriteVariableLine> setWriteVariableLines, ref string prr, ref bool foundSiteSortBin, ref bool foundXY)
        {
            if (RegEnableWords.IsMatch(line))
            {
                LineHandle.HandleEnableWords(efuseDramTable, ref scenario, ref dramType, line);
            }
            else if (RegEfuseLine.IsMatch(line) || hasReadWaferData)
            {
                LineHandle.HandleEfuseLine(efuseRows, i, line, hasReadWaferData);
                if (lines[i + 1].Contains("[INFO]"))
                {
                    hasReadWaferData = false;
                }
            }
            else if (RegReadWaferData.IsMatch(line))
            {
                hasReadWaferData = true;
                hasLimitStart = false;
            }
            else if (RegPRRCode.IsMatch(line))
            {
                prrRows.Add(new PrrLine(line, i).ToRow());
            }
            else if (hasLimitStart)
            {
                string[] arr = lines[i].Trim().Split(' ');
                if (arr[0].Contains('\t') || string.IsNullOrEmpty(arr[0])) //For the data exception handle
                {
                    return;
                }
                LineHandle.HandleLimitStart(limitRows, ref hasLimitStart, ref startIndex, lines, i, arr);
            }
            else if (RegLimite.IsMatch(line))
            {
                hasLimitStart = true;
                startIndex = i;
            }
            else if (RegSetWriteVariableLine.IsMatch(line))
            {
                LineHandle.HandleSetWriteVariableLine(setWriteVariableLines, line);
            }
            else if (RegSiteSortBin.IsMatch(line))
            {
                foundSiteSortBin = true;
                i++;
            }
            else if (RegXY.IsMatch(line))
            {
                foundXY = true;
                i++;
            }
            else if (RegPRR.IsMatch(line))
            {
                string[] array = line.Split([' ', '[', ']', '\'', '.', '\t']);
                prr = array.Last();
            }
        }

        private static void SetDiceInfo(string currentJobStage, int jobNum, string scenario, string dramType, List<EfuseRow> efuseRows, List<PrrRow> prrRows, List<LimitRow> limitRows, List<SetWriteVariableLine> setWriteVariableLines, List<DiceInfo> diceInfos, Dictionary<string, BinInfo> binInfos, Dictionary<string, CoordInfo> xyInfos)
        {
            var efuseRowsLast = efuseRows.GroupBy(x => x.BankConfig.ToLower() + "#" + x.SubConfig.ToLower() + "#" + x.Site).Select(x => x.Last()).ToList();
            var efuseRowDic = efuseRowsLast.GroupBy(x => x.Site).ToDictionary(x => x.Key, x => x.ToList());
            var prrRowDicLast = prrRows.GroupBy(x => new { x.Site, Type = x.Type?.ToUpperInvariant() }).Select(x => x.Last()).ToList();
            var prrRowDic = prrRowDicLast.GroupBy(x => x.Site).ToDictionary(x => x.Key, x => x.ToList());
            var limitRowDic = limitRows.Where(x => x != null).GroupBy(x => x.Site).ToDictionary(x => x.Key, x => x.ToList());
            foreach (KeyValuePair<int, List<EfuseRow>> siteData in efuseRowDic)
            {
                var diceInfo = new DiceInfo
                {
                    Site = siteData.Key
                };
                var xyInfo = new CoordInfo();
                bool tryGetXY = xyInfos.TryGetValue(siteData.Key.ToString(), out xyInfo);
                if (tryGetXY)
                {
                    diceInfo.XCoor = int.Parse(xyInfo!.X);
                    diceInfo.YCoor = int.Parse(xyInfo!.Y);
                }
                var binInfo = new BinInfo();
                bool tryGetBin = binInfos.TryGetValue(siteData.Key.ToString(), out binInfo);
                if (tryGetBin)
                {
                    diceInfo.Sort = int.Parse(binInfo!.Sort);
                    diceInfo.SortBin = int.Parse(binInfo!.Bin);
                }
                diceInfo.Scenario = scenario;
                diceInfo.DramType = dramType;
                diceInfo.CurrentJobStage = currentJobStage;
                diceInfo.JobNum = jobNum;
                diceInfo.EFuseLotNumber = siteData.Value.FirstOrDefault(r => r.BankConfig.EqualsIgnoreCase("Bank_ecid") && r.SubConfig.EqualsIgnoreCase("lot_id"))?.Data?.ToString();
                diceInfo.EFuseWaferId = siteData.Value.FirstOrDefault(r => r.BankConfig.EqualsIgnoreCase("Bank_ecid") && r.SubConfig.EqualsIgnoreCase("wafer_id"))?.DecValue?.ToString();
                diceInfo.EFuseDieX = siteData.Value.FirstOrDefault(r => r.BankConfig.EqualsIgnoreCase("Bank_ecid") && r.SubConfig.EqualsIgnoreCase("x_coordinate"))?.DecValue?.ToString();
                diceInfo.EFuseDieY = siteData.Value.FirstOrDefault(r => r.BankConfig.EqualsIgnoreCase("Bank_ecid") && r.SubConfig.EqualsIgnoreCase("y_coordinate"))?.DecValue?.ToString();
                diceInfo.EfuseRows = siteData.Value;
                if (prrRowDic.TryGetValue(siteData.Key, out List<PrrRow>? prrRowsForSite))
                {
                    diceInfo.PrrRows = prrRowsForSite;
                }
                if (limitRowDic.TryGetValue(siteData.Key, out List<LimitRow>? limitForSite))
                {
                    limitForSite = [.. limitForSite.GroupBy(x => x.Site + x.TestName).Select(x => x.Last())];
                    diceInfo.LimitRows = limitForSite;
                }
                diceInfo.Items = ConvertToEfuseDatalogItem(siteData.Value);
                diceInfo.SetWriteVariableLines = [.. setWriteVariableLines.Where(x => x.Site.EqualsIgnoreCase(diceInfo.Site.ToString()))];
                diceInfos.Add(diceInfo);
            }
        }

        private static Dictionary<string, EfuseDatalogItem> ConvertToEfuseDatalogItem(List<EfuseRow> efuseRows)
        {
            List<EfuseDatalogItem> items = [];
            foreach (EfuseRow efuseRow in efuseRows)
            {
                string order = "";
                var item = new EfuseDatalogItem(efuseRow.BankConfig, efuseRow.SubConfig, efuseRow.SortedBits, efuseRow.Data, order);
                items.Add(item);
            }
            return items.ToDictionary(x => x.Block + "#" + x.Id, x => x, StringExtensions.IgnoreCase);
        }
    }
}
