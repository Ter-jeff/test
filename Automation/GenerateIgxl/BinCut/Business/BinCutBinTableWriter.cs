using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutBinTableWriter
    {
        private const string Bin1 = "Bin1";
        private const string BinX = "BinX";
        private const string BinY = "BinY";

        public void GenBinTable(Dictionary<string, List<BinCutSourceItem>> sourceRowDic, BinCutInputData binCutInputManager, List<BinCutBinningItem> existHvccFlags)
        {
            BinningTable binningTable = binCutInputManager.BinningTables.First();
            BinCutFlowTable binCutFlowTable = binCutInputManager.BinCutFlowTables.First();
            var binTableBinCutSheet = new BinTableSheet("Bin_Table_BinCut");

            var tempBinItem = new List<BinTableRow>();

            tempBinItem.AddRange(GenerateBinTableRowsForBadBin(sourceRowDic, binningTable, binCutFlowTable));

            //HBV
            var list = new List<string>();
            list.AddRange(binCutInputManager.BinCutFlowTables.GetPerfromanceModeInFlow(EnumBinCutTableType.Hv));
            list = list.Distinct().ToList();
            var levelList = new List<EnumColumnName> { EnumColumnName.TD, EnumColumnName.Mbist, EnumColumnName.ILB, EnumColumnName.ELB };
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                levelList.Add(EnumColumnName.RTOS);
            }

            foreach (string mode in list)
            {
                foreach (EnumColumnName level in levelList)
                {
                    EnumColumnName domain = level;
                    string flag = "F_" + domain + "_" + mode + "_HBV";
                    string binName = "Bin_" + domain + "_" + mode + "_HBV";
                    if (!existHvccFlags.Select(x => x.FlagName).Contains(flag))
                    {
                        var tmp = new BinCutBinningItem("", "", "", "", level)
                        {
                            BinName = binName,
                            FlagName = flag,
                            PerformanceMode = mode
                        };
                        existHvccFlags.Add(tmp);
                    }
                }
            }
            binTableBinCutSheet.Rows.AddRange(GenerateBinTableRows(existHvccFlags));

            //Other
            binTableBinCutSheet.Rows.AddRange(GenerateCommonBinTable());

            //For mainEndFlowSheet
            var binTableRows = new List<BinTableRow>();

            //BV
            binTableRows.AddRange(tempBinItem);

            if (TestPlanStatic.HarvestingTruthTableSheets == null || (TestPlanStatic.HarvestingTruthTableSheets != null && !TestPlanStatic.HarvestingTruthTableSheets.Any() && string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder)))
            {
                //BinX BinY
                BinNumberSingletonLegacy.Instance().SetStartBinNumber(8000);
                binTableRows.AddRange(GeneratePowerBinningBinTableRows(sourceRowDic, binningTable));
                //Bin_BinX & Bin_BinY
                binTableRows.AddRange(GeneratePassBinXyBinTable());
            }
            binTableBinCutSheet.Rows.AddRange(binTableRows);
            AddVersion(binTableBinCutSheet, binningTable);
            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirBinCut, binTableBinCutSheet);

            #region MainInitEnableWd
            var flags = binTableBinCutSheet.Rows.SelectMany(x => x.ItemList.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, "BinCut", FolderStructure.DirMain);
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Vddbin_COF_Instance", "nop", "BinCut", FolderStructure.DirMain);
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Vddbin_COF_Instance_with_PerEqnLog", "nop", "BinCut", FolderStructure.DirMain);
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Vddbin_COF_StepInheritance", "nop", "BinCut", FolderStructure.DirMain);
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("Vddbin_CMEM_Collection", "nop", "BinCut", FolderStructure.DirMain);
            }
            #endregion
        }

        private void AddVersion(BinTableSheet binTableBinCutSheet, BinningTable binningTable)
        {
            var rows = binTableBinCutSheet.Rows.Where(x => !x.IsBackup).ToList();
            if (rows.Count > 6)
            {
                string binningRev = GenBinningTableVersion(binningTable);
                string binCutFileName = LocalSpecs.BinCutFileName == "N/A"
                    ? LocalSpecs.BinCutFileName
                    : Regex.Split(Path.GetFileNameWithoutExtension(LocalSpecs.BinCutFileName), "_Real").First() +
                      Path.GetExtension(LocalSpecs.BinCutFileName);
                string binCutPostFileName = LocalSpecs.BinCutPostFileName == "N/A"
                    ? LocalSpecs.BinCutPostFileName
                    : Regex.Split(Path.GetFileNameWithoutExtension(LocalSpecs.BinCutPostFileName), "_Real").First() +
                      Path.GetExtension(LocalSpecs.BinCutPostFileName);
                string csv = LocalSpecs.PatternListCsvFileName == "N/A"
                    ? LocalSpecs.PatternListCsvFileName
                    : Regex.Split(Path.GetFileNameWithoutExtension(LocalSpecs.PatternListCsvFileName), "_TW").First() +
                      Path.GetExtension(LocalSpecs.PatternListCsvFileName);
                string scgh = LocalSpecs.OriScghFileName == "N/A"
                    ? LocalSpecs.OriScghFileName
                    : Path.GetFileName(LocalSpecs.OriScghFileName);
                string testplan = LocalSpecs.OriTestPlanFileName == "N/A"
                    ? LocalSpecs.OriTestPlanFileName
                    : Path.GetFileName(LocalSpecs.OriTestPlanFileName);
                rows.ElementAt(0).Comment = binningRev;
                rows.ElementAt(1).Comment = binCutFileName;
                rows.ElementAt(2).Comment = binCutPostFileName;
                rows.ElementAt(3).Comment = csv;
                rows.ElementAt(4).Comment = scgh;
                rows.ElementAt(5).Comment = testplan;
            }
        }

        private List<BinTableRow> GenerateCommonBinTable()
        {
            var binTableRows = new List<BinTableRow>();

            var row = new BinTableRow { Name = "Bin_Vddbinning_IDS_fail", ItemList = "F_Vddbinning_IDS_fail" };
            row.Items.Add("T");
            row.Op = "AND";
            BinNumResult binNumResult = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", row);
            row.Bin = binNumResult.BinNumInfo.HardBin.ToString("G15");
            row.Sort = binNumResult.SoftBin.ToString("G15");
            row.Result = binNumResult.BinNumInfo.Status;
            binTableRows.Add(row);

            row = new BinTableRow
            {
                Name = BinCutConstant.VddBinningFailStop,
                ItemList = BinCutConstant.VddBinningFailStopFlag + ",Vddbinning_Fail_Stop_Flag" //Todo wait mark modify VBA
            };
            SetItemResultToTrue(ref row);
            row.Op = "OR";
            BinNumResult binNumInfoFailStop = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", row);
            row.Bin = binNumInfoFailStop.BinNumInfo.HardBin.ToString("G15");
            row.Sort = binNumInfoFailStop.SoftBin.ToString("G15");
            row.Result = binNumInfoFailStop.BinNumInfo.Status;
            binTableRows.Add(row);

            bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);
            row = new BinTableRow
            {
                Name = isCs ? BinCutConstant.VddBinningInterpolationFailCs : BinCutConstant.VddBinningInterpolationFail,
                ItemList = isCs ? BinCutConstant.VddBinningInterpolationFailFlagCs : BinCutConstant.VddBinningInterpolationFailFlag,
            };
            row.Items.Add("T");
            row.Op = "AND";
            BinNumResult binNumInfoInterpolationFail = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", row);
            row.Bin = binNumInfoInterpolationFail.BinNumInfo.HardBin.ToString("G15");
            row.Sort = binNumInfoInterpolationFail.SoftBin.ToString("G15");
            row.Result = binNumInfoInterpolationFail.BinNumInfo.Status;
            binTableRows.Add(row);

            var binTableRow = new BinTableRow { Name = BinCutConstant.PowerBinningFail, ItemList = BinCutConstant.PowerBinningFailFlag };
            binTableRow.Items.Add("T");
            binTableRow.Op = "AND";
            BinNumResult binNumInfoPowerBinning = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", binTableRow);
            binTableRow.Bin = binNumInfoPowerBinning.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Sort = binNumInfoPowerBinning.SoftBin.ToString("G15");
            binTableRow.Result = binNumInfoPowerBinning.BinNumInfo.Status;
            binTableRows.Add(binTableRow);
            return binTableRows;
        }

        internal static void SetItemResultToTrue(ref BinTableRow row)
        {
            row?.Items.AddRange(Enumerable.Repeat("T", row.ItemList.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Length));
        }

        private BinTableRow GenerateIdsBinTable(string binName)
        {
            var row = new BinTableRow
            {
                Name = Combination.CombineByUnderLine(new List<string> { "Bin", binName, "IDS" }),
                ItemList = "F_PassBinCut_1,F_PassBinCut_2,F_PassBinCut_3," + "F_IDS_" + binName
            };
            GetIdsBinItems(binName, row);

            row.Op = "AND";
            BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", row);
            row.Sort = binNumInfo.SoftBin.ToString("G15");

            row.Result = "Pass";
            row.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
            return row;
        }

        private IEnumerable<BinTableRow> GeneratePassBinXyBinTable()
        {
            var binTableRows = new List<BinTableRow>();
            var binList = new List<string> { BinX, BinY };
            foreach (string bin in binList)
            {
                var rowXy = new BinTableRow
                {
                    Name = Combination.CombineByUnderLine(new List<string> { "Bin", bin }),
                    ItemList = "F_PassBinCut_1,F_PassBinCut_2,F_PassBinCut_3",
                    Op = "AND"
                };

                if (Regex.IsMatch(bin, BinX, RegexOptions.IgnoreCase))
                {
                    rowXy.Items.AddRange(new List<string> { "F", "T", "F" });
                }
                else if (Regex.IsMatch(bin, BinY, RegexOptions.IgnoreCase))
                {
                    rowXy.Items.AddRange(new List<string> { "F", "F", "T" });
                }
                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", rowXy);
                rowXy.Sort = binNumInfo.SoftBin.ToString("G15");
                rowXy.Result = binNumInfo.BinNumInfo.Status;
                rowXy.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                binTableRows.Add(rowXy);
            }
            return binTableRows;
        }

        internal string GenBinningTableVersion(BinningTable binningTable)
        {
            string binningRev = null;
            if (binningTable != null)
            {
                if (Regex.IsMatch(binningTable.TitleRow1, "Rev:", RegexOptions.IgnoreCase))
                {
                    string[] lineSpt = binningTable.TitleRow1.Split(new[] { '\t' }, StringSplitOptions.None);
                    if (lineSpt.Length > 1)
                    {
                        binningRev = lineSpt[0].Trim() + lineSpt[1].Trim();
                        if (binningRev.Contains("."))
                        {
                            binningRev = binningRev.Replace('.', 'P');
                        }
                        else
                        {
                            binningRev += "P0";
                        }
                    }
                }
            }
            return binningRev;
        }

        private List<BinTableRow> GenerateBinTableRows(List<BinCutBinningItem> existHvccFlags)
        {
            var binTableRows = new List<BinTableRow>();
            binTableRows.AddRange(GetBinTableRows(existHvccFlags));
            return binTableRows;
        }

        private List<BinTableRow> GetBinTableRows(List<BinCutBinningItem> existHvccFlags)
        {
            var binTableRows = new List<BinTableRow>();
            foreach (BinCutBinningItem mode in existHvccFlags)
            {
                //Bin_TD_ME005_HBV =>F_TD_ME005_HBV
                var row = new BinTableRow();
                string flagName = mode.FlagName;
                string binTableName = mode.BinName;
                row.Name = binTableName;
                row.ItemList = flagName;
                row.Items.Add("T");
                row.Op = "AND";
                string category1 = mode.Level.Trim().Replace(" ", "_");
                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", category1, "", row);
                row.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                row.Sort = binNumInfo.SoftBin.ToString("G15");
                row.Result = binNumInfo.BinNumInfo.Status;
                binTableRows.Add(row);
            }
            return binTableRows;
        }

        private List<BinTableRow> GenerateBinTableRowsForBadBin(Dictionary<string, List<BinCutSourceItem>> sourceRowDic, BinningTable binningTable, BinCutFlowTable binCutFlowTable)
        {
            var binTableRows = new List<BinTableRow>();
            bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);
            var blockList = new List<string> { "Mbist", "TD", "ELB", "ILB" };
            Dictionary<string, string> pinDic = GetPinDicFromBinning(binningTable, binCutFlowTable, isCs);
            var typeList = new List<string> { "BV" };
            bool isFirst = true;
            string binningRev = GenBinningTableVersion(binningTable);
            if (!isCs)
            {
                blockList.Add("RTOS");
                typeList.Add("IDS");
            }
            foreach (string type in typeList)
            {
                foreach (string block in blockList)
                {
                    foreach (KeyValuePair<string, string> pinItem in pinDic)
                    {
                        string blockName = block;
                        if (block.Equals("RTOS", StringComparison.OrdinalIgnoreCase))
                        {
                            var sourceData = sourceRowDic.Where(x => x.Key.ContainsIgnoreCase(pinItem.Key.ToUpper())).
                                SelectMany(y => y.Value.FindAll(z => z.TargetPerformanceMode.Equals(pinItem.Key, StringComparison.OrdinalIgnoreCase)
                                    && z.Level.Equals("SPI Binning", StringComparison.OrdinalIgnoreCase))).ToList();
                            if (sourceData.Any())
                            {
                                blockName = GetTypeFromFlowName(sourceData.First().ColumnContent);
                                foreach (BinCutSourceItem data in sourceData)
                                {
                                    string tmpName = GetTypeFromFlowName(data.ColumnContent);
                                    if (!blockName.Equals(tmpName, StringComparison.OrdinalIgnoreCase) && type.Equals("BV"))
                                    {
                                        string errorMessage =
                                            $"The mode {pinItem.Key} has multiple types in FUNC block !!!";
                                        ErrorReportManager.AddError(BinCutErrorType.W_Flow_01, NeededSheets.BcFlow, sourceData.First().RowNum, 0, $"The mode {pinItem.Key} has multiple types in FUNC block !!!", new string[] { pinItem.Key });
                                        break;
                                    }
                                }
                            }
                        }
                        string mode = pinItem.Key;
                        string category1 = mode.Substring(0, mode.Length - 1);
                        string category2 = blockName + "_" + type;
                        if (type == "BV")
                        {
                            var lRow = new BinTableRow();
                            string name = Combination.CombineByUnderLine(new List<string> { mode, blockName, type }).Replace(" ", "_");
                            lRow.Name = "Bin_" + name;
                            lRow.ItemList = "F_" + name;
                            lRow.Items.Add("T");
                            lRow.Op = "AND";
                            if (isFirst)
                            {
                                lRow.Comment = binningRev;
                                isFirst = false;
                            }
                            string key = mode + "," + category2;
                            if (BinNumberSingleton.BincutBinNums.ContainsKey(key))
                            {
                                lRow.Bin = BinNumberSingleton.BincutBinNums[key].BinNumInfo.HardBin.ToString("G15");
                                lRow.Sort = BinNumberSingleton.BincutBinNums[key].SoftBin.ToString("G15");
                                lRow.Result = BinNumberSingleton.BincutBinNums[key].BinNumInfo.Status;
                            }
                            else
                            {
                                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", category1, category2, lRow);
                                lRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                                lRow.Sort = binNumInfo.SoftBin.ToString("G15");
                                lRow.Result = binNumInfo.BinNumInfo.Status;
                            }
                            binTableRows.Add(lRow);
                        }
                        else if (type == "IDS")
                        {
                            var lRow = new BinTableRow();
                            string name = Combination.CombineByUnderLine(new List<string> { mode, blockName, type }).Replace(" ", "_");
                            lRow.Name = "Bin_" + name;
                            lRow.ItemList = "F_" + name;
                            lRow.Items.Add("T");
                            lRow.Op = "AND";
                            string key = mode + "," + category2;
                            if (BinNumberSingleton.BincutBinNums.ContainsKey(key))
                            {
                                lRow.Bin = BinNumberSingleton.BincutBinNums[key].BinNumInfo.HardBin.ToString("G15");
                                lRow.Sort = BinNumberSingleton.BincutBinNums[key].SoftBin.ToString("G15");
                                lRow.Result = BinNumberSingleton.BincutBinNums[key].BinNumInfo.Status;
                            }
                            else
                            {
                                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", category1, category2, lRow);
                                lRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                                lRow.Sort = binNumInfo.SoftBin.ToString("G15");
                                lRow.Result = binNumInfo.BinNumInfo.Status;
                            }
                            binTableRows.Add(lRow);
                        }
                    }
                }
            }
            return binTableRows;
        }

        internal string GetTypeFromFlowName(string content)
        {
            if (Regex.IsMatch(content, "ELB|ILB", RegexOptions.IgnoreCase))
            {
                return "RTOS";
            }

            if (Regex.IsMatch(content, "TMPS", RegexOptions.IgnoreCase))
            {
                return "TMPS";
            }

            if (Regex.IsMatch(content, "CPM", RegexOptions.IgnoreCase))
            {
                return "CPM";
            }

            return "RTOS";
        }

        private List<BinTableRow> GeneratePowerBinningBinTableRows(Dictionary<string, List<BinCutSourceItem>> sourceRowDic, BinningTable binningTable)
        {
            var binTableRows = new List<BinTableRow>();
            var blockList = new List<string> { "Mbist", "TD", "ELB", "ILB" };
            Dictionary<string, string> pinDic = GetPinDicFromBinning(binningTable);
            var binList = new List<string> { BinX, BinY };
            var typeList = new List<string> { "BV" };
            bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);

            if (!isCs)
            {
                blockList.Add("RTOS");
                typeList.Add("IDS");
            }
            foreach (string bin in binList)
            {
                binTableRows.Add(GenerateIdsBinTable(bin));
                foreach (string type in typeList)
                {
                    foreach (string block in blockList)
                    {
                        foreach (KeyValuePair<string, string> pinItem in pinDic)
                        {
                            string blockName = block;
                            if (block.Equals("RTOS", StringComparison.OrdinalIgnoreCase))
                            {
                                var sourceData = sourceRowDic.Where(x => x.Key.ContainsIgnoreCase(pinItem.Key.ToUpper())).
                                    SelectMany(y => y.Value.FindAll(z => z.TargetPerformanceMode.Equals(pinItem.Key, StringComparison.OrdinalIgnoreCase) &&
                                        z.Level.Equals("SPI Binning", StringComparison.OrdinalIgnoreCase))).ToList();
                                if (sourceData.Any())
                                {
                                    blockName = GetTypeFromFlowName(sourceData.First().ColumnContent);
                                }
                            }

                            string pin = pinItem.Value;
                            string mode = pinItem.Key;

                            if (type == "BV")
                            {
                                var lRow = new BinTableRow();
                                string name = Combination.CombineByUnderLine(new List<string> { bin, mode, blockName, type }).Replace(" ", "_");
                                lRow.Name = "Bin_" + name;
                                lRow.ItemList = "F_" + name + ",F_PassBinCut_1,F_PassBinCut_2,F_PassBinCut_3,F_" + pin + "_" + mode + "_IDS";
                                GetBinItems(bin, lRow);
                                lRow.Items.Add("F");
                                lRow.Op = "AND";
                                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", lRow);
                                lRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                                lRow.Sort = binNumInfo.SoftBin.ToString("G15");
                                lRow.Result = binNumInfo.BinNumInfo.Status;
                                binTableRows.Add(lRow);
                            }
                            else if (type == "IDS")
                            {
                                var lRow = new BinTableRow();
                                string name = Combination.CombineByUnderLine(new List<string> { bin, mode, blockName, type }).Replace(" ", "_");
                                lRow.Name = "Bin_" + name;
                                lRow.ItemList = "F_" + name + ",F_PassBinCut_1,F_PassBinCut_2,F_PassBinCut_3,F_" + pin + "_" + mode + "_IDS";
                                GetBinItems(bin, lRow);
                                lRow.Items.Add("T");
                                lRow.Op = "AND";
                                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", "", "", lRow);
                                lRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                                lRow.Sort = binNumInfo.SoftBin.ToString("G15");
                                lRow.Result = binNumInfo.BinNumInfo.Status;
                                binTableRows.Add(lRow);
                            }
                        }
                    }
                }
            }
            return binTableRows;
        }

        internal void GetBinItems(string bin, BinTableRow row)
        {
            if (row.Name.Equals("Bin_BinX", StringComparison.OrdinalIgnoreCase))
            {
                row.Items.AddRange(new List<string>
                {
                    "F", "T", "F"
                });
            }
            else if (row.Name.Equals("Bin_BinY", StringComparison.OrdinalIgnoreCase))
            {
                row.Items.AddRange(new List<string>
                {
                    "F", "F", "T"
                });
            }
            else
            {
                if (bin == Bin1)
                {
                    row.Items.AddRange(new List<string>
                    {
                        "T", "T", "F", "F"
                    });
                }
                else if (bin == BinX)
                {
                    row.Items.AddRange(new List<string>
                    {
                        "T", "F", "T","F"
                    });
                }
                else if (bin == BinY)
                {
                    row.Items.AddRange(new List<string>
                    {
                        "T", "F", "F","T"
                    });
                }
            }
        }

        internal void GetIdsBinItems(string bin, BinTableRow row)
        {
            if (bin == Bin1)
            {
                row.Items.AddRange(new List<string>
                {
                    "T", "F", "F", "T"
                });
            }
            else if (bin == BinX)
            {
                row.Items.AddRange(new List<string>
                {
                    "F", "T", "F", "T"
                });
            }
            else if (bin == BinY)
            {
                row.Items.AddRange(new List<string>
                {
                    "F", "F", "T", "T"
                });
            }
        }

        internal Dictionary<string, string> GetPinDicFromBinning(BinningTable binningTable, BinCutFlowTable binCutFlowTable = null, bool isCs = false)
        {
            var pinDic = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
            if (binningTable == null)
            {
                return pinDic;
            }

            if (isCs)
            {
                int loopCnt = binCutFlowTable.Rows[0].PinInfos.Count;
                foreach (BinCutFlowSheetRow row in binCutFlowTable.Rows)
                {
                    int evaluateBinCount = 0;
                    for (int i = 0; i < loopCnt; i++)
                    {
                        if (row.TableType.Equals(EnumBinCutTableType.Lv) && row.PinInfos[i].PinContext.IndexOf("evaluate bin", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            evaluateBinCount++;
                        }
                        if (evaluateBinCount > 1) // multi rails
                        {
                            if (!pinDic.ContainsKey(row.PerformanceMode.Split('_').First()))
                            {
                                pinDic.Add(row.PerformanceMode.Split('_').First(), row.BinningDomain);
                            }
                            break;
                        }
                    }
                }
            }

            foreach (BinningRow row in binningTable.Rows)
            {
                if (row.RowData[binningTable.BinnedIdx].Equals("True", StringComparison.OrdinalIgnoreCase))
                {
                    if (!pinDic.ContainsKey(row.RowData[binningTable.ModeIdx]))
                    {
                        pinDic.Add(row.RowData[binningTable.ModeIdx], "VDD_" + row.RowData[binningTable.DomainIdx]);
                    }
                }

            }
            return pinDic;
        }
    }
}
