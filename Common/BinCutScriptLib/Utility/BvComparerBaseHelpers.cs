using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Flow;

namespace BinCutScriptLib.Utility
{
    internal static class BvComparerBaseHelpers
    {
        internal static BvResults GetBvResults(BvResults bvResults, Dictionary<string, double> bvs)
        {
            var result = new BvResults();
            List<string> arr = [.. BinCutConfig.SafeVoltageFollowPayload.Split(':').Last().Split(',')];
            foreach (KeyValuePair<string, double> bv in bvs)
            {
                if (bvResults.Exists(x => x.PinName == bv.Key) && arr.Exists(x => x == bv.Key))
                {
                    var bvResult = new BvResult()
                    {
                        PinName = bv.Key,
                        Voltage = bvResults.Find(x => x.PinName == bv.Key)!.Voltage
                    };
                    result.Add(bvResult);
                }
                else
                {
                    result.Add(new BvResult() { PinName = bv.Key, Voltage = 1000 * bv.Value });
                }
            }

            return result;
        }

        internal static BinCutFlowSheetRow? GetBinCutFlowSheetRow(EnumBinCutFlowType enumBinCutFlowType, string job, string bvName, EnumBinCutTableType enumBinCutTableType, EnumBinCutTableBinType enumBinCutTableBinType)
        {
            BinCutFlowSheetRow? row = null;
            if (enumBinCutFlowType == EnumBinCutFlowType.Outside)
            {
                foreach (BinCutFlowTables postFlowSheet in BinCutData.PostFlowSheets)
                {
                    row = postFlowSheet.FindRow(job, bvName, EnumBinCutTableType.Post, enumBinCutTableBinType);
                    if (row != null)
                    {
                        break;
                    }
                }
            }
            else
            {
                row = BinCutData.BinCutFlowTables.FindRow(job, bvName, enumBinCutTableType, enumBinCutTableBinType);
            }

            return row;
        }

        internal static double GetPercentageValue(string express, double retVal)
        {
            if (Reg.RegexAllRatio.IsMatch(express)) //+5.5%
            {
                _ = double.TryParse(Reg.RegexAllRatio.Match(express).Groups["ratio"].ToString().Replace(" ", ""), out double ratio);
                retVal = Math.Round(retVal * (1 + (ratio / 100)), 6, MidpointRounding.AwayFromZero);
            }

            if (Reg.RegexAllmV2.IsMatch(express))   //+5.5mV
            {
                _ = double.TryParse(Reg.RegexAllmV.Match(express).Groups["value"].ToString().Replace(" ", ""), out double value);
                retVal = Math.Round(value, 6, MidpointRounding.AwayFromZero);
            }
            else if (Reg.RegexAllmV3.IsMatch(express))   //+5.5mV
            {
                _ = double.TryParse(Reg.RegexAllmV.Match(express).Groups["value"].ToString().Replace(" ", ""), out double value);
                retVal = Math.Round(retVal + value, 6, MidpointRounding.AwayFromZero);
            }
            else if (Reg.RegexAllmV1.IsMatch(express))   //5.5mV (MS001)
            {
                _ = double.TryParse(Reg.RegexAllmV.Match("^" + express).Groups["value"].ToString().Replace(" ", ""), out double value);
                retVal = Math.Round(value, 6, MidpointRounding.AwayFromZero);
            }
            return retVal;
        }

        internal static (List<string>, string) GetFunction(BinCutExpress binCutExpress)
        {
            var expressList = new List<string>();
            string function = "";
            if (binCutExpress.IsFunction())
            {
                string parameter = Reg.RegexFunction.Match(binCutExpress.Express).Groups["parameter"].ToString();
                function = binCutExpress.Express[..binCutExpress.Express.IndexOf('(')].Trim();
                if (parameter.Contains(','))
                {
                    expressList = [.. parameter.Split(',').Select(x => x.Trim())];
                }
            }
            return (expressList, function);
        }

        internal static BinCutFlowSheetRow? GetRow(EnumBinCutFlowType enumBinCutFlowType, string job, string bvName, EnumBinCutTableType enumBinCutTableType, SiteInfo siteInfo)
        {
            EnumBinCutTableBinType tableBinType = EnumBinCutTableBinType.Bin1;
            if (siteInfo.Bin == 1)
            {
                tableBinType = EnumBinCutTableBinType.Bin1;
            }
            else if (siteInfo.Bin == 2)
            {
                tableBinType = EnumBinCutTableBinType.BinX;
            }
            else if (siteInfo.Bin == 3)
            {
                tableBinType = EnumBinCutTableBinType.BinY;
            }

            BinCutFlowSheetRow? row = null;
            if (enumBinCutFlowType == EnumBinCutFlowType.Outside)
            {
                foreach (BinCutFlowTables postflow in BinCutData.PostFlowSheets)
                {
                    row = postflow.FindRow(job, bvName, EnumBinCutTableType.Post, tableBinType);
                    if (row != null)
                    {
                        break;
                    }
                }
            }
            else
            {
                row = BinCutData.BinCutFlowTables.FindRow(job, bvName, enumBinCutTableType, tableBinType);
            }
            if (row == null)
            {
                string errorMessage = $"{bvName} can't be found in flow !!!";
                if (!ErrorReportManager.GetErrorList().Select(x => x.Message).Contains(errorMessage))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_13, "Datalog", 0, 0, $"{bvName} can't be found in flow !!!", [bvName]);
                    BinCutController.Controller.RichTextBoxAppend(errorMessage, Color.Red);
                }
            }
            return row;
        }

        internal static BinCutFlowSheetRow? GetRowCsharp(EnumBinCutFlowType enumBinCutFlowType, string job, string bvName, EnumBinCutTableType enumBinCutTableType, SiteInfo siteInfo, InstanceBinCut instanceBinCut)
        {
            EnumBinCutTableBinType tableBinType = EnumBinCutTableBinType.Bin1;
            if (siteInfo.Bin == 1)
            {
                tableBinType = EnumBinCutTableBinType.Bin1;
            }
            else if (siteInfo.Bin == 2)
            {
                tableBinType = EnumBinCutTableBinType.BinX;
            }
            else if (siteInfo.Bin == 3)
            {
                tableBinType = EnumBinCutTableBinType.BinY;
            }

            BinCutFlowSheetRow? row = GetBinCutFlowSheetRow(enumBinCutFlowType, job, bvName, enumBinCutTableType, tableBinType);

            if (row == null)
            {
                if (instanceBinCut.InstanceRow != null)
                {
                    bvName = instanceBinCut.InstanceRow.Args[0];
                    row = GetBinCutFlowSheetRow(enumBinCutFlowType, job, bvName, enumBinCutTableType, tableBinType);
                }
            }

            if (row == null)
            {
                string errorMessage = $"{bvName} can't be found in flow !!!";
                if (!ErrorReportManager.GetErrorList().Select(x => x.Message).Contains(errorMessage))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_14, "Datalog", 0, 0, $"{bvName} can't be found in flow !!!", [bvName]);
                    BinCutController.Controller.RichTextBoxAppend(errorMessage, Color.Red);
                }
            }
            return row;
        }

        internal static double GetValueByFormula(string function, List<double> doubleList)
        {
            if (function.EqualsIgnoreCase("Max"))
            {
                double value = doubleList.Max(x => x);
                return doubleList.Find(x => x == value);
            }
            if (function.EqualsIgnoreCase("Min"))
            {
                double value = doubleList.Min(x => x);
                return doubleList.Find(x => x == value);
            }
            return 0;
        }

        internal static void ReplacePin(BinCutExpress binCutExpress)
        {
            //replace VDD_CPU
            foreach (KeyValuePair<string, EnumPowerType> item in BinCutConfig.PowerType)
            {
                if (binCutExpress.Express.StartsWithIgnoreCase(item.Key))
                {
                    binCutExpress.Title = item.Key;
                    binCutExpress.Express = binCutExpress.Express.Replace(item.Key, "").Trim();
                }
            }
        }

        internal static double HandleBinProduct(BinCutExpress binCutExpress, PowerZone powerZone, double retVal)
        {
            int bin = 1;
            if (binCutExpress.Express.Contains("BINX", StringComparison.OrdinalIgnoreCase))
            {
                bin = 2;
            }
            else if (binCutExpress.Express.Contains("BINY", StringComparison.OrdinalIgnoreCase))
            {
                bin = 3;
            }

            for (int step = 0; step < powerZone.AllSteps.Count; step++)
            {
                if (powerZone.AllSteps[step].EqName == 1 && powerZone.AllSteps[step].Bin == bin)
                {
                    retVal = powerZone.AllSteps[step].BinningProduct;
                    break;
                }
            }

            return retVal;
        }

        internal static double HandleE1Product(BinCutExpress binCutExpress, SiteInfo siteInfo, PowerZone powerZone)
        {
            int bin = siteInfo.Bin;
            string[] spt = binCutExpress.Express.Split([' ']);
            //from E1 to get "1"
            int eqNo = int.Parse(spt[1][1..]);
            (bool flowControl, double value) = GetBinningProduct(powerZone, bin, eqNo);
            if (!flowControl)
            {
                return value;
            }

            return 0;
        }

        private static (bool flowControl, double value) GetBinningProduct(PowerZone powerZone, int bin, int eqNo)
        {
            for (int step = 0; step < powerZone.AllSteps.Count; step++)
            {
                if (powerZone.AllSteps[step].EqName == eqNo && powerZone.AllSteps[step].Bin == bin)
                {
                    return (flowControl: false, value: powerZone.AllSteps[step].BinningProduct);
                }
            }

            return (flowControl: true, value: default);
        }

        internal static double HandleE1Voltage(BinCutExpress binCutExpress, SiteInfo siteInfo, PowerZone powerZone)
        {
            int bin = siteInfo.Bin;
            string[] spt = binCutExpress.Express.Split([' ']);
            //from E1 to get "1"
            int eqNo = int.Parse(spt[1][1..]);
            double retVal = GetLvcc(powerZone, bin, eqNo);

            return retVal;
        }

        private static double GetLvcc(PowerZone powerZone, int bin, int eqNo)
        {
            for (int step = 0; step < powerZone.AllSteps.Count; step++)
            {
                if (powerZone.AllSteps[step].EqName == eqNo && powerZone.AllSteps[step].Bin == bin)
                {
                    return powerZone.AllSteps[step].Lvcc;
                }
            }

            return 0;
        }
    }
}
