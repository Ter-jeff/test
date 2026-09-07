using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Comparer
{
    internal class AdjustVddBinningComparer
    {
        public static void CompareAdjustVddBin(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<AdjustVddBinningRow> adjustVddBinningRows, EnumJob enumJob, bool isEqnBin, out List<int> binXList)
        {
            binXList = [];
            for (int i = 0; i < adjustVddBinningRows.Count; i++)
            {
                int site = adjustVddBinningRows[i].Site;
                string powerName = adjustVddBinningRows[i].PowerName;
                string type = adjustVddBinningRows[i].Type;
                double value = adjustVddBinningRows[i].Value;
                string unit = adjustVddBinningRows[i].Unit;
                if (type == enumJob.ToString())
                {
                    type = "CP";
                }

                for (int pwrIdx = 0; pwrIdx < siteInfoArray[site].AllPowers.Count; pwrIdx++)
                {
                    if (siteInfoArray[site].AllPowers[pwrIdx].PinMode.Contains(powerName))
                    {
                        PowerZone pwrRef = siteInfoArray[site].AllPowers[pwrIdx];
                        if (pwrRef.FinalStep == -1 || pwrRef.FinalStep == pwrRef.GetPosCount())
                        {
                            continue;
                        }
                        CompareRow(streamWriter, siteInfoArray, adjustVddBinningRows, isEqnBin, binXList, i, site, type, value, unit, pwrRef);
                        break;
                    }
                }
            }
        }

        private static void CompareRow(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<AdjustVddBinningRow> adjustVddBinningRows, bool isEqnBin, List<int> binXList, int i, int site, string type, double value, string unit, PowerZone powerZone)
        {
            double pwrLvcc = powerZone.GetFinalLvcc();
            double pwrIds = powerZone.IdsValue;
            double pwrBinNum = siteInfoArray[site].Bin;
            double pwrEqn = powerZone.PossibleSteps[powerZone.FinalStep].EqName;
            switch (type)
            {
                case "BinCut":
                    if ((int)pwrBinNum != (int)value)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrBinNum.ToString(), "<Adjust_VddBinning> : BinCut Num search");
                    }
                    break;
                case "EQN": //Compare EQN
                case "Equation_Num": //Compare EQN
                    HandleEquationNum(streamWriter, siteInfoArray, adjustVddBinningRows, isEqnBin, binXList, i, site, value, powerZone, pwrEqn);
                    break;
                case "CP": //Compare LVCC
                case "BV":
                    if (value < 5.0)
                    {
                        value *= 1000.0;
                    }

                    if (value > 1000.0)
                    {
                        if ((int)pwrLvcc != (int)value)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrLvcc.ToString(), "<Adjust_VddBinning> : LVCC search");
                        }
                    }
                    else
                    {
                        if (Math.Abs(pwrLvcc - value) > 0.001)
                        {
                            BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrLvcc.ToString(), "<Adjust_VddBinning> : LVCC search");
                        }
                    }
                    break;
                case "VDD": //Compare Product Value
                case "Product":
                    HandleProduct(streamWriter, siteInfoArray, adjustVddBinningRows, i, site, value, powerZone);
                    break;
                case "IDS": //Compare IDS Value
                    HandleIds(streamWriter, siteInfoArray, adjustVddBinningRows, isEqnBin, binXList, i, site, value, unit, powerZone, pwrIds, pwrEqn);
                    break;
                case "Monotonicity_Offset":
                    int monoResult = powerZone.MonoAdjust ? 1 : 0;
                    if (monoResult != (int)value)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), monoResult.ToString(), "<Adjust_VddBinning> : Monotonicity_Offset");
                    }
                    break;
            }
        }

        private static void HandleIds(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<AdjustVddBinningRow> adjustVddBinningRows, bool isEqnBin, List<int> binXList, int i, int site, double value, string unit, PowerZone powerZone, double pwrIds, double pwrEqn)
        {
            if (Reg.RegexUnit.IsMatch(unit))
            {
                if (unit == "u")
                {
                    value *= 1.0e-6;
                }

                if (unit == "m")
                {
                    value *= 1.0e-3;
                }
            }

            if (Math.Abs(pwrIds - value) > 0.001)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrIds.ToString(), "<Adjust_VddBinning> : IDS search");
            }
            if (isEqnBin)
            {
                string pwrModeEqn = powerZone.Mode + "_E" + (int)pwrEqn;
                BinningTable binningTable = BinCutData.BinningTables.First();
                string idsRail = binningTable.Rows.Find(x => x.RowData[binningTable.ModeEqnIdx] == pwrModeEqn)!.RowData[binningTable.BincutCalcIdSrailIdx];
                string mappingModeEqn =
                    binningTable.Rows.FindLast(
                        x => x.RowData[binningTable.DomainIdx] == idsRail)!.RowData[
                            binningTable.ModeEqnIdx];
                var idsBinXList =
                    binningTable.Rows.Select(x => new { modeEqn = x.RowData[binningTable.ModeEqnIdx], BinXIDS = double.Parse(x.RowData[binningTable.IdsMaxIdx]) }).ToList();
                if (idsBinXList.Find(x => x.modeEqn == mappingModeEqn)!.BinXIDS < value)
                {
                    if (!binXList.Contains(site))
                    {
                        binXList.Add(site);
                    }
                }
            }
        }

        private static void HandleProduct(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<AdjustVddBinningRow> adjustVddBinningRows, int i, int site, double value, PowerZone powerZone)
        {
            if (true) //5/15 add for interpolation cause no product value
            {
                double pwrProductVal = powerZone.GetFinalProductValue(streamWriter);
                if (value < 5.0)
                {
                    value *= 1000.0;
                }

                if (value > 1000.0)
                {
                    if ((int)pwrProductVal != (int)value)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrProductVal.ToString(), "<Adjust_VddBinning> : Product search");
                    }
                }
                else
                {
                    if (Math.Abs(pwrProductVal - value) > 0.001)
                    {
                        BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrProductVal.ToString(), "<Adjust_VddBinning> : Product search");
                    }
                }
            }
        }

        private static void HandleEquationNum(StreamWriter streamWriter, SiteInfo[] siteInfoArray, List<AdjustVddBinningRow> adjustVddBinningRows, bool isEqnBin, List<int> binXList, int i, int site, double value, PowerZone powerZone, double pwrEqn)
        {
            if ((int)pwrEqn != (int)value)
            {
                BinCutPrint.PrintDifference(streamWriter, siteInfoArray, adjustVddBinningRows[i].Line, site, value.ToString(), pwrEqn.ToString(), "<Adjust_VddBinning> : EQN search");
            }
            if (isEqnBin)
            {
                string pwrModeEqn = powerZone.Mode + "_E" + (int)pwrEqn;
                BinningTable binningTable = BinCutData.BinningTables.First();
                var eqnBinList =
                    binningTable.Rows.Select(x => new { modeEqn = x.RowData[binningTable.ModeEqnIdx], EqnBin = x.RowData[binningTable.EqnBinIdx] }).ToList();
                if (eqnBinList.Exists(x => x.modeEqn == pwrModeEqn && x.EqnBin.EqualsIgnoreCase("BINX")))
                {
                    if (!binXList.Contains(site))
                    {
                        binXList.Add(site);
                    }
                }
            }
        }

        public static void CompareProductIdentifier(StreamWriter streamWriter, List<ProductIdentifierLineRow> productIdentifierLineRows, List<int> binXList)
        {
            foreach (ProductIdentifierLineRow productIdentifierLineRow in productIdentifierLineRows)
            {
                if ((productIdentifierLineRow.ProductIdentifier == 0 && binXList.Contains(productIdentifierLineRow.Site))
                    || (productIdentifierLineRow.ProductIdentifier == 1 && !binXList.Contains(productIdentifierLineRow.Site)))
                {
                    BinCutPrint.PrintAdjustProductIdentifier(streamWriter, productIdentifierLineRow, binXList);
                }
            }
        }
    }
}
