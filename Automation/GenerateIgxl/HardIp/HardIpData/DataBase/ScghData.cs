using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Rtos.Input;

using Automation.Reader.ScghFile.ProCharPatternSet.Base;
using Automation.Reader.ScghFile.ProCharPatternSet.Business;
using Automation.Singleton;

using CommonLib.Extension;

using OfficeOpenXml;

using ScghLib.Reader;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class ScghData
    {
        #region Field

        private readonly List<HardIpScghSheet> _hardipSheetList;                                                                         //original Hardip scgh sheet List
        public List<ProdCharSheetRow> ScanScghRowList = new List<ProdCharSheetRow>();                    //orignal scan scgh sheet list
        public List<ProdCharSheetRow> HardIpSheetRowList = new List<ProdCharSheetRow>();
        public List<ProdCharSheetRow> ConvertedPatternRowListByHardip = new List<ProdCharSheetRow>();
        public List<ProdCharSheetRow> ConvertedPatternRowListByAll = new List<ProdCharSheetRow>();
        public static Dictionary<string, ProdCharSheetRow> AliasMappingTable = new Dictionary<string, ProdCharSheetRow>();
        #endregion

        #region Properity

        public List<ProdCharSheetRow> GetBistPatterns
        {
            get
            {
                var sheetList = new List<string> {
                    NeededSheets.MbistCharScgCpu,
                NeededSheets.MbistCharScgGpu,
                NeededSheets.MbistCharScgSoc,
                NeededSheets.MbistCharScg};
                return ConvertedPatternRowListByAll.Where(p => sheetList.Contains(p.SourceSheetName)).ToList();
            }
        }

        public List<ProdCharSheetRow> GetScanPatterns
        {
            get
            {
                var sheetList = new List<string>
                {
                    NeededSheets.ScanScghCpu,
                    NeededSheets.ScanScghGpu,
                    NeededSheets.ScanScghSoc,
                    NeededSheets.ScanScgh
                };
                return ConvertedPatternRowListByAll.Where(p => sheetList.Contains(p.SourceSheetName)).ToList();
            }
        }

        #endregion

        public ScghData()
        {
            _hardipSheetList = new List<HardIpScghSheet>();
            HardIpSheetRowList = new List<ProdCharSheetRow>();
            ScanScghRowList = new List<ProdCharSheetRow>();
            AliasMappingTable.Clear();
        }

        #region Static Function

        public ScghData LoadEfuseFromHardIpBistScghData(ExcelWorkbook scghWorkbook, bool isOtp = false)
        {
            List<HardIpScghSheet> hardIpScghSheets = GetEfuseFromHardIpScghSheets(scghWorkbook, isOtp);
            List<ProdCharSheet> prodCharSheets = GetEfuseProdCharSheets(scghWorkbook, isOtp);
            var scgh = new ScghData();
            foreach (HardIpScghSheet hardIpScghSheet in hardIpScghSheets)
            {
                scgh.AddHardipSheet(hardIpScghSheet);
            }

            foreach (ProdCharSheet prodCharSheet in prodCharSheets)
            {
                scgh.AddNonHardipSheet(prodCharSheet);
            }

            AliasMappingTable = scgh.HardIpSheetRowList.GroupBy(p => p.Item).ToDictionary(p => p.Key, p => p.ToList()[0]);

            #region Hardip
            var patset = new HardIpPatSetConstructor(scgh.HardIpSheetRowList, null as List<string>);
            List<ProdCharPatternSetHardIp> result = patset.WorkFlow();
            scgh.ConvertedPatternRowListByHardip = new List<ProdCharSheetRow>();
            foreach (ProdCharPatternSetHardIp row in result)
            {
                var tempRow = (ProdCharSheetRow)row.ProdCharRow;
                ProdCharSheetRow scghRow = tempRow.Copy();
                scghRow.InitList = row.InitList.Values.Select(p => p.PatternName).ToList();
                scghRow.PayloadList = row.PayloadList.Select(x => x.PatternName).ToList();
                scgh.ConvertedPatternRowListByHardip.Add(scghRow);
            }

            var removedScghItems = new List<ProdCharSheetRow>();
            //remove used items or init pattern in SCGH
            foreach (ProdCharSheetRow row in scgh.ConvertedPatternRowListByHardip)
            {
                if (IsInitPattern(row.PayloadValue))
                {
                    removedScghItems.Add(row);
                }
                else if (!IsHasInit(row) && IsItemUsed(scgh.ConvertedPatternRowListByHardip, row))
                {
                    removedScghItems.Add(row);
                }
            }

            foreach (ProdCharSheetRow removeitem in removedScghItems)
            {
                scgh.ConvertedPatternRowListByHardip.Remove(removeitem);
            }

            #endregion
            return scgh;
        }

        public static ScghData LoadScghData(ExcelWorkbook scghWorkbook, bool isConvertedAll = false, bool isRemoveEfuse = true)
        {
            List<HardIpScghSheet> hardIpScghSheets = GetHardIpScghSheets(scghWorkbook, isRemoveEfuse);
            List<ProdCharSheet> prodCharSheets = GetProdCharSheets(scghWorkbook, isRemoveEfuse);

            var scgh = new ScghData();
            foreach (HardIpScghSheet hardIpScghSheet in hardIpScghSheets)
            {
                scgh.AddHardipSheet(hardIpScghSheet);
            }

            foreach (ProdCharSheet prodCharSheet in prodCharSheets)
            {
                scgh.AddNonHardipSheet(prodCharSheet);
            }

            AliasMappingTable = scgh.HardIpSheetRowList.GroupBy(p => p.Item).ToDictionary(p => p.Key, p => p.ToList()[0]);

            #region Hardip
            var patset = new HardIpPatSetConstructor(scgh.HardIpSheetRowList, null as List<string>);
            List<ProdCharPatternSetHardIp> result = patset.WorkFlow();
            scgh.ConvertedPatternRowListByHardip = new List<ProdCharSheetRow>();
            //scgh.ConvertedPatternRowListByHardip.AddRange(initRowList);
            foreach (ProdCharPatternSetHardIp row in result)
            {
                var tempRow = (ProdCharSheetRow)row.ProdCharRow;
                ProdCharSheetRow scghRow = tempRow.Copy();
                scghRow.InitList = row.InitList.Values.Select(p => p.PatternName).ToList();
                scghRow.PayloadList = row.PayloadList.Select(x => x.PatternName).ToList();
                scgh.ConvertedPatternRowListByHardip.Add(scghRow);
            }

            var removedScghItems = new List<ProdCharSheetRow>();
            //remove used items or init pattern in SCGH
            foreach (ProdCharSheetRow row in scgh.ConvertedPatternRowListByHardip)
            {
                if (IsInitPattern(row.PayloadValue))
                {
                    removedScghItems.Add(row);
                }
                else if (!IsHasInit(row) && IsItemUsed(scgh.ConvertedPatternRowListByHardip, row))
                {
                    removedScghItems.Add(row);
                }
            }

            foreach (ProdCharSheetRow removeitem in removedScghItems)
            {
                scgh.ConvertedPatternRowListByHardip.Remove(removeitem);
            }

            #endregion

            #region All
            if (isConvertedAll)
            {
                var total = new List<ProdCharSheetRow>();
                total.AddRange(scgh.HardIpSheetRowList);
                total.AddRange(scgh.ScanScghRowList);
                var patsetAll = new HardIpPatSetConstructor(total, (List<string>)null);
                List<ProdCharPatternSetHardIp> resultAll = patsetAll.WorkFlow();
                scgh.ConvertedPatternRowListByAll = new List<ProdCharSheetRow>();
                foreach (ProdCharPatternSetHardIp row in resultAll)
                {
                    var tempRow = (ProdCharSheetRow)row.ProdCharRow;
                    ProdCharSheetRow scghRow = tempRow.Copy();
                    scghRow.InitList = row.InitList.Values.Select(p => p.PatternName).ToList();
                    scghRow.PayloadList = row.PayloadList.Select(x => x.PatternName).ToList();
                    scgh.ConvertedPatternRowListByAll.Add(scghRow);
                }

                var removedScghItemsAll = new List<ProdCharSheetRow>();
                //remove used items or init pattern in SCGH
                foreach (ProdCharSheetRow row in scgh.ConvertedPatternRowListByAll)
                {
                    if (IsInitPattern(row.PayloadValue))
                    {
                        removedScghItemsAll.Add(row);
                    }
                    else if (!IsHasInit(row) && IsItemUsed(scgh.ConvertedPatternRowListByAll, row))
                    {
                        removedScghItemsAll.Add(row);
                    }
                }

                foreach (ProdCharSheetRow removeitem in removedScghItemsAll)
                {
                    scgh.ConvertedPatternRowListByAll.Remove(removeitem);
                }
            }
            #endregion

            return scgh;
        }

        private static List<ProdCharSheet> GetProdCharSheets(ExcelWorkbook scghWorkbook, bool isRemoveEfuse)
        {
            var nonHardipSheetNameList = new List<string>
            {
                NeededSheets.ScanScghCpu,
                NeededSheets.ScanScghGpu,
                NeededSheets.ScanScghSoc,
                NeededSheets.ScanScgh,
                NeededSheets.MbistCharScgCpu,
                NeededSheets.MbistCharScgGpu,
                NeededSheets.MbistCharScgSoc,
                NeededSheets.MbistCharScg,
                NeededSheets.SpiScghCpu,
                NeededSheets.SpiScghGpu,
                NeededSheets.SpiScghSoc
            };

            var prodCharSheets = new List<ProdCharSheet>();
            if (scghWorkbook == null)
            {
                return prodCharSheets;
            }

            foreach (string sheetName in nonHardipSheetNameList)
            {
                ExcelWorksheet worksheet = scghWorkbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    continue;
                }

                ProdCharSheetReader nonHardipReader = null;
                if (sheetName == NeededSheets.SpiScghCpu || sheetName == NeededSheets.SpiScghGpu ||
                    sheetName == NeededSheets.SpiScghSoc)
                {
                    var formater = new InputFormat(worksheet);
                    int format = formater.GetFormatNumber();
                    if (format == 0)
                    {
                        nonHardipReader = new ProdCharSheetReader();
                    }
                }
                else
                {
                    nonHardipReader = new ProdCharSheetReader();
                }

                if (nonHardipReader == null)
                {
                    continue;
                }

                ProdCharSheet scanSheet = nonHardipReader.ReadSheet(worksheet);
                scanSheet.RowList = RemoveUnValidScghRows(scanSheet.RowList, isRemoveEfuse);
                prodCharSheets.Add(scanSheet);
            }
            return prodCharSheets;
        }

        private static List<HardIpScghSheet> GetHardIpScghSheets(ExcelWorkbook scghWorkbook, bool isRemoveEfuse)
        {
            var hardipSheetNameList = new List<string>
            {
                NeededSheets.HardIpScgh,
                NeededSheets.HardIpScghC,
                NeededSheets.HardIpScghG,
                NeededSheets.HardIpScghS,
                NeededSheets.PmicScgh
            };
            var hardIpScghSheets = new List<HardIpScghSheet>();
            foreach (string sheetName in hardipSheetNameList)
            {
                ExcelWorksheet worksheet = scghWorkbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    continue;
                }

                var hardipReader = new HardIpProdCharReader();
                HardIpScghSheet hardipSheet = hardipReader.ReadSheet(worksheet);
                hardipSheet.RowList = RemoveUnValidScghRows(hardipSheet.RowList, isRemoveEfuse);
                hardIpScghSheets.Add(hardipSheet);
            }
            return hardIpScghSheets;
        }

        internal List<HardIpScghSheet> GetEfuseFromHardIpScghSheets(ExcelWorkbook scghWorkbook, bool isOtp = false)
        {
            var hardipSheetNameList = new List<string>
            {
                NeededSheets.HardIpScgh,
                NeededSheets.HardIpScghC,
                NeededSheets.HardIpScghG,
                NeededSheets.HardIpScghS,
                NeededSheets.PmicScgh
            };
            var hardIpScghSheets = new List<HardIpScghSheet>();
            foreach (string sheetName in hardipSheetNameList)
            {
                ExcelWorksheet worksheet = scghWorkbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    continue;
                }

                var hardipReader = new HardIpProdCharReader();
                HardIpScghSheet hardipSheet = hardipReader.ReadSheet(worksheet);
                hardipSheet.RowList = isOtp ? GetOtpScghRows(hardipSheet.RowList) : GetEfuseScghRows(hardipSheet.RowList);
                hardIpScghSheets.Add(hardipSheet);
            }
            return hardIpScghSheets;
        }

        private static List<ProdCharSheet> GetEfuseProdCharSheets(ExcelWorkbook scghWorkbook, bool isOtp = false)
        {
            var nonHardipSheetNameList = new List<string>
            {
                NeededSheets.MbistCharScgCpu,
                NeededSheets.MbistCharScgGpu,
                NeededSheets.MbistCharScgSoc,
                NeededSheets.MbistCharScg
            };

            var prodCharSheets = new List<ProdCharSheet>();
            foreach (string sheetName in nonHardipSheetNameList)
            {
                ExcelWorksheet worksheet = scghWorkbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    continue;
                }

                ProdCharSheetReader nonHardipReader = null;
                if (sheetName == NeededSheets.SpiScghCpu || sheetName == NeededSheets.SpiScghGpu ||
                    sheetName == NeededSheets.SpiScghSoc)
                {
                    var formater = new InputFormat(worksheet);
                    int format = formater.GetFormatNumber();
                    if (format == 0)
                    {
                        nonHardipReader = new ProdCharSheetReader();
                    }
                }
                else
                {
                    nonHardipReader = new ProdCharSheetReader();
                }

                if (nonHardipReader == null)
                {
                    continue;
                }

                ProdCharSheet bistSheet = nonHardipReader.ReadSheet(worksheet);
                List<ProdCharSheetRow> fusePayloads = isOtp ? GetOtpScghRows(bistSheet.RowList) : GetEfuseScghRows(bistSheet.RowList);
                GetEfuseInit(ref fusePayloads, bistSheet.RowList);
                bistSheet.RowList = fusePayloads;
                prodCharSheets.Add(bistSheet);
            }
            return prodCharSheets;
        }

        private static void GetEfuseInit(ref List<ProdCharSheetRow> fusePls, List<ProdCharSheetRow> allBistRows)
        {
            foreach (ProdCharSheetRow fusePlItem in fusePls)
            {
                if (fusePlItem.InitList.Any())
                {
                    fusePlItem.GetInitList();
                    var realInitList = new List<string>();
                    foreach (string init in fusePlItem.InitList)
                    {
                        if (init.Equals("n/a", StringComparison.CurrentCultureIgnoreCase))
                        {
                            continue;
                        }

                        ProdCharSheetRow realInitPat = allBistRows.Find(x => !string.IsNullOrEmpty(x.Item) && x.Item.Equals(init, StringComparison.CurrentCultureIgnoreCase)) ??
                                                       allBistRows.Find(x => !string.IsNullOrEmpty(x.Mode) && x.Mode.Equals(init, StringComparison.CurrentCultureIgnoreCase));
                        if (realInitPat != null)
                        {
                            realInitList.Add(realInitPat.PayloadValue);
                        }
                    }
                    if (realInitList.Any())
                    {
                        fusePlItem.InitList = realInitList;
                    }
                }
            }
        }

        private static bool IsItemUsed(List<ProdCharSheetRow> allitems, ProdCharSheetRow row)
        {
            if (allitems.Exists(p => p.InitList.Contains(row.Item)))
            {
                return true;
            }

            return false;
        }

        private static bool IsHasInit(ProdCharSheetRow scgh)
        {
            if (scgh.InitList.Count(x => x != "" && x.ToUpper() != "N/A" && x.ToUpper() != "NA") > 0)
            {
                return true;
            }

            return false;
        }

        /*
         * Use position 3 to determine is init pattern or not.
         * e.g. PP_PTCA0_S_IN05_BI_BIST_PFF_JTG_UNS_ALLFRV_SI_SERVER_SETUP
         */
        private static bool IsInitPattern(string payload)
        {
            if (string.IsNullOrEmpty(payload))
            {
                return false;
            }

            string[] arrayList = payload.Split('_');
            return arrayList.Length > 4 && Regex.IsMatch(arrayList[3], @"_in\w+_", RegexOptions.IgnoreCase);
        }
        #endregion

        #region Member Function
        public void AddHardipSheet(HardIpScghSheet sheet)
        {
            _hardipSheetList.Add(sheet);
            HardIpSheetRowList.AddRange(sheet.RowList);
        }

        public void AddNonHardipSheet(ProdCharSheet sheet)
        {
            ScanScghRowList.AddRange(sheet.RowList);
        }

        public static List<ProdCharSheetRow> GetEfuseScghRows(List<ProdCharSheetRow> hardIpScghRowlst)
        {
            hardIpScghRowlst = hardIpScghRowlst.FindAll(x => x.PayloadValue.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).Length > 5 &&
                                                             x.PayloadValue.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries)[4].Equals("EF", StringComparison.CurrentCultureIgnoreCase) && x.Usage.Trim().Equals("1"));
            return hardIpScghRowlst;
        }

        public static List<ProdCharSheetRow> GetOtpScghRows(List<ProdCharSheetRow> hardIpScghRowlst)
        {
            hardIpScghRowlst = hardIpScghRowlst.FindAll(x => x.PayloadValue.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).Length > 5 && x.PayloadValue.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).Last().ToUpper().Contains("OTP") && x.Usage.Trim().Equals("1") && x.Block.ContainsIgnoreCase("OTP"));
            return hardIpScghRowlst;
        }

        public static List<ProdCharSheetRow> RemoveUnValidScghRows(List<ProdCharSheetRow> hardIpScghRowlst, bool isRemoveEfuse)
        {
            hardIpScghRowlst = isRemoveEfuse ? hardIpScghRowlst.FindAll(a =>
                a.Block.ToLower() != "efuse" &&
                (a.Application.ToLower() == "production" || (a.Usage != "0" && a.Application.ToUpper() != "HTOL"))
            ) : hardIpScghRowlst;
            return hardIpScghRowlst;
        }

        public ProdCharSheetRow GetProdCharSheetRow(List<string> patList)
        {
            ProdCharSheetRow targetRow = null;
            foreach (HardIpScghSheet hardipSheet in _hardipSheetList)
            {
                foreach (ProdCharSheetRow row in hardipSheet.RowList)
                {
                    if (IsMatchWithTargetRow(patList, row))
                    {
                        targetRow = row;
                        break;
                    }
                }
            }

            return targetRow;
        }

        public ProdCharSheetRow GetProdCharSheetRow(string pattern)
        {
            ProdCharSheetRow targetRow = null;
            if (pattern.Contains(";"))
            {
                string init = pattern.Split(';')[0];
                string payload = pattern.Split(';')[1];
                List<string> initList = init.Split(',').ToList();
                List<string> payloadList = payload.Split(',').ToList();
                //init1 , ...;payload
                foreach (HardIpScghSheet hardipSheet in _hardipSheetList)
                {
                    foreach (ProdCharSheetRow row in hardipSheet.RowList)
                    {
                        if (IsMatchWithTargetRow(initList, payloadList, row))
                        {
                            targetRow = row;
                            break;
                        }
                    }
                }
            }
            else
            {
                //only payload
                foreach (HardIpScghSheet hardipSheet in _hardipSheetList)
                {
                    foreach (ProdCharSheetRow row in hardipSheet.RowList)
                    {
                        if (row.PayloadValue.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                        {
                            targetRow = row;
                            break;
                        }
                    }
                }
            }
            return targetRow;
        }

        private bool IsMatchWithTargetRow(List<string> initList, List<string> payloadList, ProdCharSheetRow row)
        {
            for (int i = 0; i < initList.Count; i++)
            {
                if (row.InitList.Count <= i)
                {
                    return false;
                }

                if (!initList[i].Equals(row.InitList[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            for (int i = 0; i < payloadList.Count; i++)
            {
                if (row.PayloadList.Count <= i)
                {
                    return false;
                }

                if (!payloadList[i].Equals(row.PayloadList[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        private bool IsMatchWithTargetRow(List<string> patList, ProdCharSheetRow row)
        {
            if (patList.Count != row.InitList.Count + row.PayloadList.Count)
            {
                return false;
            }

            for (int i = 0; i < patList.Count; i++)
            {
                var total = new List<string>();
                total.AddRange(row.InitList);
                total.AddRange(row.PayloadList);
                if (!patList[i].Equals(total[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        public bool GetPatternListFromScgh(string item, ref List<List<string>> rowList)
        {
            rowList = FindAllInMode(item, ConvertedPatternRowListByHardip);
            if (rowList.Count != 0)
            {
                return true;
            }

            rowList = FindAllInItem(item, ConvertedPatternRowListByHardip);
            if (rowList.Count != 0)
            {
                return true;
            }

            rowList = FindAllInMode(item, ScanScghRowList);
            if (rowList.Count != 0)
            {
                return true;
            }

            rowList = FindAllInItem(item, ScanScghRowList);
            if (rowList.Count != 0)
            {
                return true;
            }

            rowList = FindAllInMode(item, HardIpSheetRowList);
            if (rowList.Count != 0)
            {
                return true;
            }

            rowList = FindAllInItem(item, HardIpSheetRowList);
            if (rowList.Count != 0)
            {
                return true;
            }

            rowList = new List<List<string>> { new List<string> { item } };
            return false;
        }

        protected List<List<string>> FindAllInMode(string init, List<ProdCharSheetRow> rowList)
        {
            var patternlist = new List<List<string>>();
            List<ProdCharSheetRow> rows = rowList.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase));
            foreach (ProdCharSheetRow row in rows)
            {
                var list = new List<string>();
                list.AddRange(row.GetInitList());
                list.AddRange(row.GetPayloadList());
                patternlist.Add(list);
            }
            return patternlist;
        }

        protected List<List<string>> FindAllInItem(string init, List<ProdCharSheetRow> rowList)
        {
            var patternlist = new List<List<string>>();
            List<ProdCharSheetRow> rows = rowList.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase));
            foreach (ProdCharSheetRow row in rows)
            {
                var list = new List<string>();
                list.AddRange(row.GetInitList());
                list.AddRange(row.GetPayloadList());
                patternlist.Add(list);
            }
            return patternlist;
        }

        public static List<string> GetPerformanceMode(ExcelWorkbook scghWorkbook, bool isConvertedAll = false, bool isRemoveEfuse = true)
        {
            var modes = new List<string>();
            List<ProdCharSheet> prodCharSheets = GetProdCharSheets(scghWorkbook, isRemoveEfuse);
            foreach (ProdCharSheet prodCharSheet in prodCharSheets)
            {
                List<List<string>> patternLists = prodCharSheet.GetPatternLists();
                foreach (List<string> patternList in patternLists)
                {
                    foreach (string pattern in patternList)
                    {
                        string[] arr = pattern.Split('_');
                        if (arr.Length < 10)
                        {
                            continue;
                        }

                        string mode = arr[9];
                        if (Regex.IsMatch(mode, PerformanceModeSingleton.RegContainPerformanceModeByPattern, RegexOptions.IgnoreCase))
                        {
                            if (!modes.Exists(x => x.Equals(mode, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                modes.Add(mode);
                            }
                        }
                    }
                }
            }

            modes = modes.Where(x => !(x.EndsWith("00", StringComparison.CurrentCultureIgnoreCase) ||
                x.EndsWith("XX", StringComparison.CurrentCultureIgnoreCase)))
                .Select(x => x.ToUpper()).OrderBy(x => x).Distinct().ToList();

            return modes;
        }
        #endregion
    }
}
