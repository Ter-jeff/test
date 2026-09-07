using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace TestPlanLib.BinCut.Flow
{
    public partial class BinCutFlowTables : List<BinCutFlowTable>
    {
        [GeneratedRegex(@"^M{1}[A-Z]+\w{3}$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public BinCutFlowTables(IEnumerable<BinCutFlowTable> binCutFlowTables) : base(binCutFlowTables)
        {
        }

        public BinCutFlowTables()
        {
        }

        public BinCutFlowSheetRow? GetRow(string job, string bvName, EnumBinCutTableType enumBinCutTableType, EnumBinCutTableBinType enumBinCutTableBinType)
        {
            return FindRow(job, bvName, enumBinCutTableType, enumBinCutTableBinType);
        }

        public BinCutFlowSheetRow? FindRow(string job, string bvName, EnumBinCutTableType enumBinCutTableType, EnumBinCutTableBinType enumBinCutTableBinType)
        {
            if (!Exists(x => x.FinalJob.Contains(job)) && !job.Contains("T0TX"))
            {
                return null;
            }

            BinCutFlowTable? sheet = Find(x => x.FinalJob.Contains(job));
            if (job.Contains("T0TX"))
            {
                sheet = Find(x => x.JobName.Contains(job));
            }
            if (sheet == null)
            {
                return null;
            }
            List<BinCutFlowSheetRow> rows = sheet.Rows.Exists(x => x.TableBinType == enumBinCutTableBinType && x.TableType == enumBinCutTableType) ?
                [.. sheet.Rows.Where(x => x.TableBinType == enumBinCutTableBinType && x.TableType == enumBinCutTableType)] :
                [.. sheet.Rows.Where(x => x.TableBinType == EnumBinCutTableBinType.Bin1 && x.TableType == enumBinCutTableType)];

            string mode = GetModeByName(bvName).ToUpper();
            string modeName = bvName[bvName.IndexOf(mode, StringComparison.Ordinal)..].ToUpper();

            return !rows.Exists(x => x.PerformanceMode.EqualsIgnoreCase(modeName))
                ? null
                : rows.Find(x => x.PerformanceMode.EqualsIgnoreCase(modeName));
        }

        private static string GetModeByName(string name)
        {
            //eg. VDD_SOC_MS001 => MS001
            //eg. BV_VDD_SOC_MS001 => MS001
            //eg. MC602 E1 Voltage => MC602
            string pPowerToken = "";
            string[] spt = name.Split(['_', ' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries);
            foreach (string tok in spt)
            {
                if (_regex.IsMatch(tok))
                {
                    pPowerToken = tok;
                    break;
                }
            }
            return pPowerToken;
        }

        public Dictionary<string, string> GetModeVsPowerName()
        {
            var powerDic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            for (int col = 0; col < this[0].Rows.Count; col++)
            {
                BinCutFlowSheetRow row = this[0].Rows[col];
                foreach (PinInfo voltage in row.PinInfos)
                {
                    string title = voltage.PinName.Split(',')[0].Trim();
                    string express = voltage.PinContext;
                    string mode = GetModeByName(express);
                    if (express.Contains("EVALUATE", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!powerDic.ContainsKey(mode.ToUpper()))
                        {
                            powerDic.Add(mode.ToUpper(), title.ToUpper());
                        }
                    }
                }
            }
            return powerDic;
        }

        public List<string> GetTitle(EnumJob enumJob)
        {
            BinCutFlowTable? table = Find(x => x.FinalJob.Contains(enumJob.ToString()));
            List<string> mainpins = table?.PowerPins.Keys.ToList() ?? [];
            List<string> affiliatedPin = table?.AffiliatedPin.Keys.ToList() ?? [];
            return [.. mainpins, .. affiliatedPin];
        }

        public List<Tuple<string, string>> GetSkipPwrList()
        {
            const string interpolateDonotest = "_DONOTEST";
            const string interpolateDonottest = "_DONOTTEST";
            var skipList = new List<Tuple<string, string>>();
            foreach (BinCutFlowTable tbl in this)
            {
                foreach (BinCutFlowSheetRow row in tbl.Rows)
                {
                    if ((string.IsNullOrEmpty(row.Atpg) || row.Atpg.EndsWithIgnoreCase(interpolateDonotest) || row.Atpg.EndsWithIgnoreCase(interpolateDonottest)) &&
                        (string.IsNullOrEmpty(row.Mbist) || row.Mbist.EndsWithIgnoreCase(interpolateDonotest) || row.Mbist.EndsWithIgnoreCase(interpolateDonottest)) &&
                        (string.IsNullOrEmpty(row.SpiRtos) || row.SpiRtos.EndsWithIgnoreCase(interpolateDonotest) || row.SpiRtos.EndsWithIgnoreCase(interpolateDonottest)))
                    {
                        foreach (string job in row.Job)
                        {
                            var str = new Tuple<string, string>(job, row.PerformanceMode);
                            skipList.Add(str);
                        }
                    }
                }
            }
            return skipList;
        }

        public List<string> GetPerfromanceModeInFlow(EnumBinCutTableType enumBinCutTableType)
        {
            var hvPmodes = new List<string>();
            if (enumBinCutTableType == EnumBinCutTableType.Post)
            {
                foreach (BinCutFlowTable flowTable in this)
                {
                    hvPmodes.AddRange(flowTable.Rows.Where(x => x.TableType == enumBinCutTableType).Select(x => x.PerformanceMode));
                }
            }
            else
            {
                foreach (BinCutFlowTable flowTable in this)
                {
                    hvPmodes.AddRange(flowTable.Rows.Where(x => x.TableType == enumBinCutTableType).Select(x => x.PerformanceMode.Split('_').First()));
                }
            }

            return [.. hvPmodes.Distinct()];
        }

        #region Check
        public void Check(BinningTables? binningTables)
        {
            var modes = new List<string>();
            bool needCheckMode = false;
            if (binningTables?.Count > 0)
            {
                needCheckMode = true;
                modes = [.. binningTables[0].Rows.Select(x => x.RowData[binningTables[0].ModeIdx]).Distinct()];
            }

            var evalatePins = new List<string>();
            Dictionary<string, int> powerPins = this[0].PowerPins;
            for (int i = 0; i < Count; i++)
            {
                #region syntax of mode
                if (modes.Count != 0)
                {
                    for (int j = i + 1; j < this[i].Rows.Count; j++)
                    {
                        BinCutFlowSheetRow row = this[i].Rows[j];
                        string name = row.PerformanceMode.Split('_').First();
                        if (!modes.Exists(x => x.EqualsIgnoreCase(name)))
                        {
                            string errorMessage = $"Please check the {row.PerformanceMode} syntax of performance, that should be existed in binning sheeet !!!";
                            this[i].AddError(BinCutErrorType.W_Flow_02, this[i].SheetName, row.RowNum, this[i].Indices.PerformanceModeIndex, $"Please check the {row.PerformanceMode} syntax of erformance, that should be existed in binning sheeet !!!", [row.PerformanceMode]);
                        }
                    }
                }
                #endregion

                #region headers
                if (this[i].PowerPins.Count == powerPins.Count)
                {
                    for (int j = 0; j < this[i].PowerPins.Count; j++)
                    {
                        if (!this[i].PowerPins.ElementAt(j).Key.EqualsIgnoreCase(powerPins.ElementAt(j).Key))
                        {
                            string errorMessage = $"Please check core power pin in different jobs {this[0].JobName} {powerPins.ElementAt(j).Key} vs {this[i].JobName} {this[i].PowerPins.ElementAt(j).Key} !!!";
                            this[i].AddError(BinCutErrorType.E_Flow_01, this[i].SheetName, 0, 0, $"Please check core power pin in different jobs {this[0].JobName} {powerPins.ElementAt(j).Key} vs {this[i].JobName} {this[i].PowerPins.ElementAt(j).Key} !!!", [this[0].JobName, powerPins.ElementAt(j).Key, this[i].JobName, this[i].PowerPins.ElementAt(j).Key]);
                        }
                    }
                }
                else
                {
                    string errorMessage = $"The core power count in {this[0].JobName} does not match with {this[i].JobName} !!!";
                    this[i].AddError(BinCutErrorType.E_Flow_02, this[i].SheetName, 0, 0, $"The core power count in {this[0].JobName} does not match with {this[i].JobName} !!!", [this[0].JobName, this[i].JobName]);
                }
                #endregion

                #region Check empty
                foreach (BinCutFlowSheetRow row in this[i].Rows)
                {
                    if (string.IsNullOrEmpty(row.AllOther))
                    {
                        string errorMessage = $"AllOther in row {row.RowNum} is empty !!!";
                        this[i].AddError(BinCutErrorType.W_Flow_03, this[i].SheetName, row.RowNum, this[i].Indices.AllOtherIndex, $"AllOther in row {row.RowNum} is empty !!!", [row.RowNum.ToString()]);
                    }
                    int idx = 0;
                    foreach (PinInfo vol in row.PinInfos)
                    {
                        idx++;
                        if (string.IsNullOrEmpty(vol.PinContext))
                        {
                            string errorMessage = $"{vol.PinName} in row {row.RowNum} is empty !!!";
                            this[i].AddError(BinCutErrorType.W_Flow_04, this[i].SheetName, row.RowNum, this[i].Indices.PerformanceModeIndex + idx, $"{vol.PinName} in row {row.RowNum} is empty !!!", [vol.PinName, row.RowNum.ToString()]);
                        }
                    }
                }
                #endregion

                #region check evalate by jobs
                //var currentEvalatePins = this[i].GetEvaluatePins();
                //foreach (var currentEvalatePin in currentEvalatePins)
                //{
                //    if (evalatePins.Exists(x => x.Equals(currentEvalatePin, StringComparison.CurrentCultureIgnoreCase)))
                //    {
                //        var errorMessage = string.Format("Can not BinCut search in different jobs for Pin {0} !!!",currentEvalatePin);
                //        this[i].AddError(BinCutErrorType.E_Flow, ErrorLevel.Error, this[i].SheetName, 0, 0, errorMessage);
                //    }
                //}
                //evalatePins.AddRange(currentEvalatePins);
                //evalatePins = evalatePins.Distinct().ToList();
                #endregion

                if (needCheckMode)
                {
                    CheckMode(this[i], binningTables!);
                }
            }
        }

        #region Check with other input
        public static void CheckMode(BinCutFlowTable binCutFlowTable, BinningTables binningTables)
        {
            int rowIndex = 0;
            foreach (BinCutFlowSheetRow perRow in binCutFlowTable.Rows)
            {
                rowIndex++;
                int colIndex = 0;
                foreach (PinInfo powerData in perRow.PinInfos)
                {
                    colIndex++;
                    string pmode = GetModeByName(powerData.PinContext);
                    if (!string.IsNullOrEmpty(pmode))
                    {
                        if (powerData.PinContext.Contains("Evaluate Bin"))
                        {
                            if (!perRow.PerformanceMode.Contains(pmode))
                            {
                                //string errorMessage = $"The mode {pmode} in flow can not be found in the column performance mode !!!";
                                binCutFlowTable.AddError(BinCutErrorType.E_Flow_03, binCutFlowTable.SheetName, binCutFlowTable.Indices.StartRowIndex + rowIndex, binCutFlowTable.Indices.PerformanceModeIndex + colIndex, $"The mode {pmode} in flow can not be found in the column performance mode !!!", [pmode]);
                            }
                        }
                        else
                        {
                            if (!CheckFlowSheetPmode(pmode, binningTables))
                            {
                                //string errorMessage = $"The mode {pmode} in flow can not be found in binning, binning_binX ro binning_binY!!!";
                                binCutFlowTable.AddError(BinCutErrorType.W_Flow_05, binCutFlowTable.SheetName, binCutFlowTable.Indices.StartRowIndex + rowIndex, binCutFlowTable.Indices.PerformanceModeIndex + colIndex, $"The mode {pmode} in flow can not be found in binning, binning_binX ro binning_binY!!!", [pmode]);
                            }
                        }
                    }
                }
            }
        }

        public void CheckFlowSheet(Dictionary<string, string> idsList)
        {
            foreach (BinCutFlowTable table in this)
            {
                BinCutFlowSheetRow row = table.Rows[0];
                int cnt = 0;
                foreach (PinInfo voltage in row.PinInfos)
                {
                    string idsName = "";
                    bool flag = false;
                    bool flagContains = false;
                    cnt++;
                    foreach (KeyValuePair<string, string> item in idsList)
                    {
                        if (item.Key.EqualsIgnoreCase("IDS_" + voltage.PinName))
                        {
                            flag = true;
                            idsName = item.Key;
                            break;
                        }

                        if (item.Key.Contains(voltage.PinName, StringComparison.OrdinalIgnoreCase))
                        {
                            flagContains = true;
                            idsName = item.Key;
                            break;
                        }
                    }

                    if (!flag)
                    {
                        if (flagContains)
                        {
                            //string errorMessage = $"The ids name of {voltage.PinName} in flow can not be found in efuseBitDef, and only found relative name {idsName} !!!";
                            table.AddError(BinCutErrorType.W_efuseBitDef_02, table.SheetName, table.Indices.StartRowIndex, table.Indices.PerformanceModeIndex + cnt, $"The ids name of {voltage.PinName} in flow can not be found in efuseBitDef, and only found relative name {idsName} !!!", [voltage.PinName, idsName]);
                        }
                        else
                        {
                            //string errorMessage = $"The ids name of {voltage.PinName} in flow can not be found in efuseBitDef !!!";
                            table.AddError(BinCutErrorType.W_efuseBitDef_03, table.SheetName, table.Indices.StartRowIndex, table.Indices.PerformanceModeIndex + cnt, $"The ids name of {voltage.PinName} in flow can not be found in efuseBitDef !!!", [voltage.PinName]);
                        }
                    }
                }
            }
        }

        private static bool CheckFlowSheetPmode(string pmode, BinningTables binningTables)
        {
            var modes = new List<string>();
            for (int idx = 0; idx < binningTables.Count; idx++)
            {
                BinningTable vddBinTable = binningTables[idx];
                List<string> modeList = vddBinTable.Rows.ConvertAll(y => y.RowData[vddBinTable.ModeIdx].ToUpper());
                modes.AddRange(modeList);
            }
            modes = [.. modes.Distinct()];

            return modes.Exists(x => x == pmode);
        }
        #endregion
        #endregion
    }

    public class PinInfo
    {
        public string PinName = "";
        public string PinContext = "";
    }

    public class SelsramInfo
    {
        public string PinName = "";
        public string Bit = "";
    }
}
