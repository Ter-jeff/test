using System;
using System.Collections.Generic;
using System.IO;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib.Comparer.BVChecker.GetVoltage
{
    public abstract class GetVoltageBase(EnumJob enumJob, InstanceBinCut instanceBinCut)
    {
        public InstanceBinCut InstanceBinCut = instanceBinCut;
        public EnumJob Job = enumJob;
        public int TestCat = (int)enumJob;

        private static readonly Dictionary<(int JobId, string Type), Func<BinningTable, int>> _indexSelectors = new()
        {
            { (0, "TD"), bt => bt.OffsetCp1TdIdx },
            { (0, "MBIST"), bt => bt.OffsetCp1BistIdx },
            { (0, "FUNC"), bt => bt.OffsetCp1FuncIdx },
            { (1, "TD"), bt => bt.OffsetCp2TdIdx },
            { (1, "MBIST"), bt => bt.OffsetCp2BistIdx },
            { (1, "FUNC"), bt => bt.OffsetCp2FuncIdx },
            { (2, "TD"), bt => bt.OffsetFt1TdIdx },
            { (2, "MBIST"), bt => bt.OffsetFt1BistIdx },
            { (2, "FUNC"), bt => bt.OffsetFt1FuncIdx },
            { (3, "TD"), bt => bt.OffsetFt2TdIdx },
            { (3, "MBIST"), bt => bt.OffsetFt2BistIdx },
            { (3, "FUNC"), bt => bt.OffsetFt2FuncIdx },
            { (4, "TD"), bt => bt.OffsetQaTdIdx },
            { (4, "MBIST"), bt => bt.OffsetQaBistIdx },
            { (4, "FUNC"), bt => bt.OffsetQaFuncIdx }
        };

        private static (bool flowControl, double value) HandleJob(BinningTable binningTable, string currentInstType, int rowIdx, int jobId)
        {
            string typeUpper = currentInstType?.ToUpperInvariant() ?? string.Empty;

            if (!_indexSelectors.TryGetValue((jobId, typeUpper), out Func<BinningTable, int>? selector))
            {
                return (flowControl: true, value: default);
            }

            int targetIdx = selector(binningTable);

            if (targetIdx == -1)
            {
                return (flowControl: false, value: 0);
            }

            List<string> rowData = binningTable.Rows[rowIdx].RowData;
            if (targetIdx >= 0 && targetIdx < rowData.Count && double.TryParse(rowData[targetIdx], out double value))
            {
                return (flowControl: false, value);
            }

            return (flowControl: false, value: 0);
        }

        private static (bool flowControl, double value) HandleJob0(BinningTable binningTable, string t, int r) => HandleJob(binningTable, t, r, 0);

        private static (bool flowControl, double value) HandleJob1(BinningTable binningTable, string t, int r) => HandleJob(binningTable, t, r, 1);

        private static (bool flowControl, double value) HandleJob2(BinningTable binningTable, string t, int r) => HandleJob(binningTable, t, r, 2);

        private static (bool flowControl, double value) HandleJob3(BinningTable binningTable, string t, int r) => HandleJob(binningTable, t, r, 3);

        private static (bool flowControl, double value) HandleJob4(BinningTable binningTable, string t, int r) => HandleJob(binningTable, t, r, 4);

        public abstract double GetValue(BinCutExpress binCutExpress, SiteInfo siteInfo, BvName bvName, StreamWriter streamWriter);

        protected double GetRamValue(BinCutExpress binCutExpress, BinningTable binningTable)
        {
            double retVal;
            string mode = BinCutAlgorithmService.GetModeByName(binCutExpress.Express).Length == 0 ? binCutExpress.Title : BinCutAlgorithmService.GetModeByName(binCutExpress.Express);
            int rowIdx = BinCutAlgorithmService.GetOtherIndex(mode, binningTable);
            if (rowIdx == -1)
            {
            }
            int colunmIndex = binningTable.GetColumnIndex(binCutExpress.Express);

            if (binCutExpress.Express.Contains("PRODUCT", StringComparison.OrdinalIgnoreCase))
            {
                double dGbTmp = 0.0;
                retVal = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CIdx]) +
                      double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CpGbIdx]);

                if (Reg.RegexAllProduct.IsMatch(binCutExpress.Express) && binCutExpress.Express.Contains("GB", StringComparison.OrdinalIgnoreCase))
                {
                    int lvVddBinGbIdx = GetLvVddBinGbIdx(binningTable);
                    dGbTmp = double.Parse(binningTable.Rows[rowIdx].RowData[lvVddBinGbIdx]);
                }
                retVal -= dGbTmp;
            }
            else if (binCutExpress.Express.Contains("LVCC", StringComparison.OrdinalIgnoreCase))
            {
                double dProductTmp = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CIdx]) +
                                     double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CpGbIdx]);

                int lvOthRaiGbIdx = GetLvVddBinGbIdx(binningTable);
                double dGbTmp = double.Parse(binningTable.Rows[rowIdx].RowData[lvOthRaiGbIdx]);
                retVal = dProductTmp - dGbTmp;
            }
            else if (binCutExpress.Express.Contains("HVCC", StringComparison.OrdinalIgnoreCase))
            {
                int hvOthRalIdx = GetHvVddBinIdx(binningTable);
                retVal = double.Parse(binningTable.Rows[rowIdx].RowData[hvOthRalIdx]);
            }
            else if (colunmIndex != -1)
            {
                retVal = double.Parse(binningTable.Rows[rowIdx].RowData[colunmIndex]);
            }
            else
            {
                double dProductTmp = double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CIdx]) +
                                     double.Parse(binningTable.Rows[rowIdx].RowData[binningTable.CpGbIdx]);

                int lvOthRaiGbIdx = GetLvVddBinGbIdx(binningTable);
                double dGbTmp = double.Parse(binningTable.Rows[rowIdx].RowData[lvOthRaiGbIdx]);
                retVal = dProductTmp - dGbTmp;
            }
            return retVal;
        }

        protected int GetLvVddBinGbIdx(BinningTable binningTable)
        {
            int lvVddBinGbIdx = -1;
            //T0TX_HOT & T0TX_ROOM are not formal job, so can't get GbIdx from job type.
            if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
            {
                TestCat = BinCutConfig.FlagT0TxHotFormat ? 5 : 6;
            }

            switch (TestCat)
            {
                case 0:
                    lvVddBinGbIdx = binningTable.CpGbIdx;
                    break;
                case 1:
                    lvVddBinGbIdx = binningTable.Cp2GbIdx;
                    break;
                case 2:
                    lvVddBinGbIdx = binningTable.Ft1GbIdx;
                    break;
                case 3:
                    lvVddBinGbIdx = binningTable.Ft2GbIdx;
                    break;
                case 4:
                    lvVddBinGbIdx = binningTable.QaGbIdx;
                    break;
                case 5:
                    lvVddBinGbIdx = binningTable.HtolGbHotIdx;
                    break;
                case 6:
                    lvVddBinGbIdx = binningTable.HtolGbRoomIdx;
                    break;
            }
            return lvVddBinGbIdx;
        }

        protected int GetHvVddBinIdx(BinningTable binningTable)
        {
            int hvVddBinIdx = -1;
            switch (TestCat)
            {
                case 0:
                    hvVddBinIdx = binningTable.CpHvIdx;
                    break;
                case 1:
                    hvVddBinIdx = binningTable.CpHvIdx;
                    break;
                case 2:
                    hvVddBinIdx = binningTable.FtHvIdx;
                    break;
                case 3:
                    hvVddBinIdx = binningTable.FtHvIdx;
                    break;
                case 4:
                    hvVddBinIdx = binningTable.QaHvIdx;
                    break;
            }
            return hvVddBinIdx;
        }

        protected static double GetE1Voltage(BinCutExpress binCutExpress, SiteInfo siteInfo, PowerZone powerZone, int bin)
        {
            string[] spt = binCutExpress.Express.Split([' ']);
            //from E1 to get "1"
            if (int.TryParse(spt[1][1..], out int eqNo))
            {
                for (int step = 0; step < powerZone.AllSteps.Count; step++)
                {
                    if (powerZone.AllSteps[step].EqName == eqNo && powerZone.AllSteps[step].Bin == bin)
                    {
                        return powerZone.AllSteps[step].Lvcc;
                    }
                }
            }
            return 0;
        }

        protected static double GetE1ProductVoltage(BinCutExpress binCutExpress, SiteInfo siteInfo, PowerZone powerZone, int bin)
        {
            double retVal = 0;
            string[] spt = binCutExpress.Express.Split([' ']);
            //from E1 to get "1"
            int eqNo = binCutExpress.ExpressType == EnumExpressType.E1Product ? int.Parse(spt[1][1..]) : 1;
            for (int step = 0; step < powerZone.AllSteps.Count; step++)
            {
                if (powerZone.AllSteps[step].EqName == eqNo && powerZone.AllSteps[step].Bin == bin)
                {
                    retVal = powerZone.AllSteps[step].BinningProduct;
                }
            }
            return retVal;
        }

        protected static double GetDynamicOffSetValue(string powerMode, int eqNo, BinningTable binningTable, int job, string currentInstType)
        {
            for (int rowIdx = 0; rowIdx < binningTable.Rows.Count; rowIdx++)
            {
                string modeInBin = binningTable.Rows[rowIdx].RowData[binningTable.ModeIdx];
                int eqnInBin = int.Parse(binningTable.Rows[rowIdx].RowData[binningTable.EqnIdx][1..]);
                if ((powerMode == modeInBin) & (eqNo == eqnInBin))
                {
                    switch (job)
                    {
                        case 0:
                            (bool flowControl, double value) = HandleJob0(binningTable, currentInstType, rowIdx);
                            if (!flowControl)
                            {
                                return value;
                            }
                            break;

                        case 1:
                            (bool flowControl1, double value1) = HandleJob1(binningTable, currentInstType, rowIdx);
                            if (!flowControl1)
                            {
                                return value1;
                            }
                            break;

                        case 2:
                            (bool flowControl2, double value2) = HandleJob2(binningTable, currentInstType, rowIdx);
                            if (!flowControl2)
                            {
                                return value2;
                            }
                            break;

                        case 3:
                            (bool flowControl3, double value3) = HandleJob3(binningTable, currentInstType, rowIdx);
                            if (!flowControl3)
                            {
                                return value3;
                            }
                            break;

                        case 4:
                            (bool flowControl4, double value4) = HandleJob4(binningTable, currentInstType, rowIdx);
                            if (!flowControl4)
                            {
                                return value4;
                            }
                            break;

                        default:
                            return 0;
                    }
                }
            }
            return 0;
        }
    }
}
