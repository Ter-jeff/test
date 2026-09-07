using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.Flow
{
    public partial class BinCutFlowTable : MySheet
    {
        //CP1
        //LV - CORE_POWER : (1)MS001 E1 Voltage (2)MS001 Evaluate Bin (3)MS001 Bin Result (4)HVCC Level at MC607
        //LV - RAM_POWER  : (1)MC604 CSRAM V (2)any string with MC604 (3)MC601 CSRAM Product (4)MC601 CSRAM Product +5%
        //LV - Others     : (1)CP LVCC (2)Product (3)800mV
        //HV - CORE_POWER : (1)HVCC Level at MC702 (2)MC606 Product +3% (3)MC606 Product (4)MS001 E1 Voltage (5)MS001 Bin Result
        //HV - RAM_POWER  : (1)MC606 CSRAM Product (2)MC606 CSRAM HVCC (3)MC601 CSRAM Product +5%
        //HV - Others     : (1)CP HVCC (2)Product (3)800mV (4)Product +5% (5)CP HVCC

        //CP2
        //LV - CORE_POWER : (1)MC701 product-CP2GB (2)MG007 product +25mV
        //LV - RAM_POWER  : (1)MC604 CSRAM V (2)any string with MC604 (3)MC601 CSRAM Product (4)MC601 CSRAM Product +5%
        //LV - Others     : (1)CP LVCC (2)Product (3)800mV
        //HV - CORE_POWER : (1)HVCC Level at MC702 (2)MC606 Product +3% (3)MC606 Product (4)MC701 product-CP2GB
        //HV - RAM_POWER  : (1)MC606 CSRAM Product (2)MC606 CSRAM HVCC (3)MC601 CSRAM Product +5%
        //HV - Others     : (1)CP HVCC (2)Product (3)800mV (4)Product +5%

        public const string RegexPerformance = "(?<pmode>M[a-zA-Z0-9]{4}[a-zA-Z0-9]?)";

        #region Core Power
        //LV CORE_POWER => MS001 E1 Voltage
        public const string RegexLvCoreVoltage = RegexPerformance + @"\s+E\d+\s+Voltage(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //LV CORE_POWER => MS001 Evaluate Bin
        public const string RegexLvCoreEvaluate = RegexPerformance + @"\s+Evaluate\s+Bin(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //LV CORE_POWER => MS001 Bin Result
        public const string RegexLvCoreResult = RegexPerformance + @"\s+Bin\s+Result(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //LV,HV CORE_POWER => HVCC Level at MC607
        public const string RegexCoreHvcc = @"HVCC\s+Level\s+at\s+" + RegexPerformance + @"(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //CORE_POWER => MC606 Product
        public const string RegexCoreProduct1 = RegexPerformance + @"\s+Product(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //CORE_POWER => MC606 E1 Product
        public const string RegexCoreE1Product = RegexPerformance + @"\s+E\d\s+Product(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //CORE_POWER => MC606 E1 Product - CP_GB_HOT
        public const string RegexCoreE1ProductGb = RegexPerformance + @"\s+E\d\s+Product\s*[+|-]\s*\w*GB\w*(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        ////CORE_POWER => MC606 E1 Product - CP_GB_HOT
        //public static string RegexCoreE1ProductGb = RegexPerformance + @"\s+E\d\s+Product\s*[+|-]\s*\w*+GB\w*+(\s*[+|-])(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //MAX(MD003 Product +10%, VDD_SRAM_SOC CP HVCC)
        public const string RegexFunction = @".*[\(](?<parameter>.*)[\)]";
        #endregion

        #region RAM Power
        //RAM_POWER => MC601 CSRAM Product
        public const string RegexRam1 = RegexPerformance + @"\s+(.|\n)*SRAM(\s+(?<ratio>[+|-]s*\d+(\.\d+)?)%)?(\s+(?<value>[+|-]s*\d+(\.\d+)?)s*mV)?";
        #endregion

        #region others
        //LV Others => CP LVCC
        public const string RegexLvOTher1 = @"\s+LVCC";
        //HV Others => CP HVCC
        public const string RegexHvOTher1 = @"\s+HVCC";
        //MPS001 CPVmax
        public const string RegexCpVmax = RegexPerformance + @"\s+CPVmax";
        //MPS001 BinningVmax
        public const string RegexBinningVmax = RegexPerformance + @"\s+BinningVmax";
        //MPS001 CPHV
        public const string RegexCphv = RegexPerformance + @"\s+CPHV";
        //MPS001 Product
        public const string RegexProduct = RegexPerformance + @"\s+Product";
        #endregion

        #region All
        //All => +5.5%
        public const string RegexAllRatio = @"(?<ratio>[+|-]\s*\d+(\.\d+)?)%";
        //All => +700.25mV 
        public const string RegexAllmV = @"(?<value>[+|-]?\s*\d+(\.\d+)?)\s*mV";
        //All => +700.25mV (MS001)
        public const string RegexAllmVWithMode = @"(?<value>[+|-]?\s*\d+(\.\d+)?)\s*mV\s*\(" + RegexPerformance + @"\)";
        //All => Product
        public const string RegexAllProduct = "Product";
        //All => Product+10.5%
        public const string RegexProductRatioMv = @"^Product\s*(" + RegexAllRatio + ")?(" + RegexAllmV + ")?";
        //All => 700.25mV (MS001 Product)
        public const string RegexmVWithProductMode = @"^(?<value>\d+(\.\d+)?)\s*mV\s*\(\s*" + RegexPerformance + @"\s*Product\s*\)$";
        //All => 700.25mV (MS001 BinSearch)
        public const string RegexmVWithBinSearchMode = @"^(?<value>\d+(\.\d+)?)\s*mV\s*\(\s*" + RegexPerformance + @"\s*BinSearch\s*\)$";
        //All => MC606 BinX Product+10%
        public const string RegexModeBinProduct = RegexPerformance + @"\s+Bin(1|X|Y)\s+Product(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        //All => MC606 BinX Product - CP_GB_HOT
        public const string RegexModeBinProductGb = RegexPerformance + @"\s+Bin(1|X|Y)\s+Product\s*[+|-]\s*\w*GB\w*(\s*(?<ratio>[+|-]\s*\d+(\.\d+)?)%)?(\s*(?<value>[+|-]\s*\d+(\.\d+)?)\s*mV)?$";
        #endregion

        [GeneratedRegex(" Bin Result", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(RegexPerformance, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("^Bincut_", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(@"^(LV|HV|NV)\s+Levels", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex(@"\s{2,}", RegexOptions.Compiled)]
        private static partial Regex MyRegex4();
        [GeneratedRegex(RegexCoreProduct1, RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex MyRegex9();

        private readonly BinCutFlowExpressionValidator _validator = new();

        public BinCutFlowColumnIndices Indices = new();

        //FT_HOT @ 85'C, mV
        public string JobName = "";
        //FT2
        public List<string> FinalJob = [];

        public Dictionary<string, int> PowerPins = [];
        public Dictionary<string, int> AffiliatedPin = [];
        public List<BinCutFlowSheetRow> Rows = [];

        public void SortByBinCutOrder(List<BinCutOrderRow> binCutOrderRows)
        {
            var lvRows = Rows.Where(x => x.TableType == EnumBinCutTableType.Lv).ToList();
            var hvRows = Rows.Where(x => x.TableType == EnumBinCutTableType.Hv).ToList();
            var postRows = Rows.Where(x => x.TableType == EnumBinCutTableType.Post).ToList();
            var lvFlowRows = new List<BinCutFlowSheetRow>();
            var lvOrders = binCutOrderRows.Where(x => x.Bincut.EqualsIgnoreCase("Search")).ToList();
            foreach (BinCutOrderRow lvOrder in lvOrders)
            {
                foreach (BinCutFlowSheetRow row in lvRows)
                {
                    if (lvOrder.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode))
                    {
                        lvFlowRows.Add(row);
                    }
                }
            }

            //remained
            foreach (BinCutFlowSheetRow row in lvRows)
            {
                if (!lvFlowRows.Exists(x => x.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode)))
                {
                    row.Nop = true;
                    lvFlowRows.Add(row);
                }
            }
            Rows = lvFlowRows;

            var hvGroups = hvRows.GroupBy(x => x.TableBinType).ToList();
            foreach (IGrouping<EnumBinCutTableBinType, BinCutFlowSheetRow> hvGroup in hvGroups)
            {
                var hvFlowRows = new List<BinCutFlowSheetRow>();
                if (hvGroup != null && hvGroup.Any())
                {
                    List<BinCutOrderRow> hvOrders = BinCutFlowTableHelpers.GetHvBinCutOrderRows(binCutOrderRows, hvGroup.First().TableBinType);
                    foreach (BinCutOrderRow hvOrder in hvOrders)
                    {
                        foreach (BinCutFlowSheetRow row in hvGroup)
                        {
                            if (hvOrder.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode))
                            {
                                hvFlowRows.Add(row);
                            }
                        }
                    }

                    foreach (BinCutFlowSheetRow row in hvGroup)
                    {
                        if (!hvFlowRows.Exists(x => x.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode)))
                        {
                            var all = hvGroup.Where(x => x.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode)).ToList();
                            foreach (BinCutFlowSheetRow one in all)
                            {
                                one.Nop = true;
                                hvFlowRows.Add(one);
                            }
                        }
                    }
                }
                Rows.AddRange(hvFlowRows);
            }

            var postFlowRows = new List<BinCutFlowSheetRow>();
            var postOrders = binCutOrderRows.Where(x => x.Bincut.EqualsIgnoreCase("Post")).ToList();
            foreach (BinCutOrderRow postOrder in postOrders)
            {
                foreach (BinCutFlowSheetRow row in postRows)
                {
                    if (postOrder.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode))
                    {
                        postFlowRows.Add(row);
                    }
                }
            }

            //remained
            foreach (BinCutFlowSheetRow row in postRows)
            {
                if (!postFlowRows.Exists(x => x.PerformanceMode.EqualsIgnoreCase(row.PerformanceMode)))
                {
                    row.Nop = true;
                    postFlowRows.Add(row);
                }
            }
            Rows.AddRange(postFlowRows);
        }

        public void Check(List<string>? binningTitleList = null, Dictionary<string, List<string>>? domainDic = null)
        {
            #region Check Syntax
            _validator.RefreshGradeSearchPins(Rows);
            foreach (BinCutFlowSheetRow row in Rows)
            {
                foreach (PinInfo voltage in row.PinInfos)
                {
                    if (voltage.PinContext.Count(x => x.Equals('(')) != voltage.PinContext.Count(x => x.Equals(')')))
                    {
                        string errorMessage = $"Please check syntax for\"{voltage.PinContext}\" ,the count of \"(\" and \")\" are mismatch !!!";
                        AddError(BinCutErrorType.E_FormatError_06, SheetName, row.RowNum, PowerPins[voltage.PinName], $"Please check syntax for\"{voltage.PinContext}\" ,the count of \"(\" and \")\" are mismatch !!!", [voltage.PinContext]);
                    }

                    string domain = voltage.PinName.Replace("VDD_", "").Split([',', ' '], StringSplitOptions.RemoveEmptyEntries).First();
                    if (!IsValidExpress(voltage.PinContext, row, domain, binningTitleList, domainDic) && !string.IsNullOrEmpty(row.PerformanceMode))
                    {
                        string errorMessage = $"The syntax \"{voltage.PinContext}\"  of {voltage.PinName} was unknown !!!";
                        AddError(BinCutErrorType.E_FormatError_07, SheetName, row.RowNum, PowerPins[voltage.PinName], $"The syntax \"{voltage.PinContext}\"  of {voltage.PinName} was unknown !!!", [voltage.PinContext, voltage.PinName]);
                    }
                }
            }
            #endregion

            #region Check Bin Result & Evaluate
            int loopCnt = Rows[0].PinInfos.Count;
            for (int j = 0; j < loopCnt; j++)
            {
                var evaluatebinMode = new List<string>();
                foreach (BinCutFlowSheetRow row in Rows)
                {
                    PinInfo rowData = row.PinInfos[j];
                    if (PowerPins.ContainsKey(rowData.PinName) && row.TableType.Equals(EnumBinCutTableType.Lv))
                    {
                        if (MyRegex1().IsMatch(rowData.PinContext.Split(' ').First()) &&
                            rowData.PinContext.Contains("evaluate bin", StringComparison.OrdinalIgnoreCase))
                        {
                            string mode = rowData.PinContext.Split(' ').First();
                            evaluatebinMode.Add(mode);
                        }
                        else
                        {
                            if (MyRegex().IsMatch(rowData.PinContext))
                            {
                                string mode = rowData.PinContext.Split(' ').First();
                                if (!evaluatebinMode.Any(x => x.EqualsIgnoreCase(mode)))
                                {
                                    string errorMessage = $"{rowData.PinContext} has not evaluate bin before Bin Result !!!";
                                    AddError(BinCutErrorType.E_FormatError_08, SheetName, row.RowNum, Indices.PerformanceModeIndex + 1 + j, $"{rowData.PinContext} has not evaluate bin before Bin Result !!!", [rowData.PinContext]);
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            #region Check BinCut DC category
            foreach (BinCutFlowSheetRow row in Rows)
            {
                if (!string.IsNullOrEmpty(row.AllOther) && !MyRegex2().IsMatch(row.AllOther) && !MyRegex3().IsMatch(row.AllOther))
                {
                    if (!MyRegex9().IsMatch(row.AllOther))
                    {
                        string errorMessage = $"The {row.AllOther} is incorrect DC category in the Bincut !!!";
                        AddError(BinCutErrorType.E_FormatError_09, SheetName, row.RowNum, Indices.AllOtherIndex, $"The {row.AllOther} is incorrect DC category in the Bincut !!!", [row.AllOther]);
                    }
                }
            }
            #endregion

            #region Check flow name
            //string errorMessages = "There are extra spaces !!!";
            foreach (BinCutFlowSheetRow row in Rows)
            {
                List<string> td = [.. row.Atpg.Split(';')];
                if (td.Any(x => MyRegex4().IsMatch(x.Trim())))
                {
                    AddError(BinCutErrorType.E_FormatError_10, SheetName, row.RowNum, Indices.AtpgIndex, "There are extra spaces !!!");
                }

                List<string> bist = [.. row.Mbist.Split(';')];
                if (bist.Any(x => MyRegex4().IsMatch(x.Trim())))
                {
                    AddError(BinCutErrorType.E_FormatError_11, SheetName, row.RowNum, Indices.MbistIndex, "There are extra spaces !!!");
                }

                List<string> func = [.. row.SpiRtos.Split(';')];
                if (func.Any(x => MyRegex4().IsMatch(x.Trim())))
                {
                    AddError(BinCutErrorType.E_FormatError_12, SheetName, row.RowNum, Indices.SpiRtosIndex, "There are extra spaces !!!");
                }
            }
            #endregion
        }

        public bool IsValidExpress(string express, BinCutFlowSheetRow binCutFlowSheetRow, string binningDomain, List<string>? headerList = null, Dictionary<string, List<string>>? domainDic = null)
        {
            return _validator.IsValidExpress(express, binningDomain, headerList, domainDic);
        }

        public List<string> GetEvaluateModes()
        {
            List<string> evaluateModes = Rows.FindAll(x => x.ExistEvaluateBin()).ConvertAll(y => y.PerformanceMode);
            return evaluateModes;
        }
    }

    public class BinCutFlowSheetRow : MyRow
    {
        #region Properity
        public bool Nop;
        public EnumBinCutTableType TableType;
        public EnumBinCutTableBinType TableBinType;
        public string BinningDomain { get; set; }
        public string PerformanceMode { get; set; }
        public List<PinInfo> PinInfos { get; set; } = [];
        public string AllOther { get; set; }
        public string Atpg { get; set; }
        public string Mbist { get; set; }
        public string SpiRtos { get; set; }
        public bool CanBeMergedAtpg { get; set; }
        public bool CanBeMergedMbist { get; set; }
        public bool CanBeMergedSpiRtos { get; set; }
        public bool Static { get; set; }
        public List<SelsramInfo> SelsramInfos { get; set; } = [];

        public List<string> Job;
        #endregion

        #region Constructor
        public BinCutFlowSheetRow(string sheetName, List<string> job)
        {
            SheetName = sheetName;
            Job = job;
            BinningDomain = "";
            PerformanceMode = "";
            AllOther = "";
            Atpg = "";
            Mbist = "";
            SpiRtos = "";
        }
        #endregion

        public bool ExistEvaluateBin()
        {
            return PinInfos.ConvertAll(x => x.PinContext.ToLower()).Exists(x => x.Contains("evaluate bin"));
        }
    }
}
