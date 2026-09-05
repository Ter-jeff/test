using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using CommonReaderLib;

using IgxlLib.Enums;

using TestPlanLib.BinCut.BinCutConfig;

namespace TestPlanLib.BinCut.Binning
{
    public partial class BinningTable : MySheet
    {
        [GeneratedRegex("Max PV", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("Min PV", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"[-|\s|+]", RegexOptions.Compiled)]
        private static partial Regex MyRegex2();
        [GeneratedRegex(@"(?<pmode>M[a-zA-Z]*)(?<modenumber>[\d]*)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex3();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();
        private static readonly Regex _regex4 = MyRegex3();

        public string Version = "";
        public string TitleRow1 = "";
        public string TitleRow2 = "";
        public string TitleRow3 = "";
        public string Job = "";

        public int BinnedIdx = -1;

        //Doamin
        public int StartRowIdx = -1;
        //Doamin: CPU/GPU/SOC
        public int DomainIdx = -1;
        //ID
        public int IdIdx = -1;
        //Mode: MC601/MC602/MC603
        public int ModeIdx = -1;
        //EQN_Bin
        public int EqnBinIdx = -1;
        public int BinXIdsMaxIdx = -1;

        //CPGB
        public int CpGbIdx = -1;
        //CP2GB
        public int Cp2GbIdx = -1;
        //FT1GB
        public int Ft1GbIdx = -1;
        //FT2GB
        public int Ft2GbIdx = -1;
        //ATE_FQAGB?
        public int QaGbIdx = -1;
        public int HtolGbIdx = -1;
        public int HtolGbRoomIdx = -1;
        public int HtolGbHotIdx = -1;

        //EQN
        public int EqnIdx = -1;
        //Mode_EQN
        public int ModeEqnIdx = -1;
        //C
        public int CIdx = -1;
        //M
        public int MIdx = -1;
        //CPIDSMax
        public int IdsMaxIdx = -1;
        //FTIDS
        public int FtMaxIdx = -1;
        //CPVmax
        public int CpVMaxIdx = -1;
        //CPVmin
        public int CpVMinIdx = -1;

        //CPHV
        public int CpHvIdx = -1;
        //FTHV
        public int FtHvIdx = -1;
        //QAHV
        public int QaHvIdx = -1;

        //SRAMthresh_Product
        public int SramthreshProductIdx = -1;
        //SRAMthresh_CP1
        public int SramthreshCp1Idx = -1;
        //SRAMthresh_CP2
        public int SramthreshCp2Idx = -1;

        //Int_Mode_L
        public int IntModeLIdx = -1;
        //Int_Mode_H
        public int IntModeHIdx = -1;
        //Int_MF
        public int IntMfIdx = -1;
        //Int_Offset
        public int IntOffsetIdx = -1;
        //Int_SkipTest
        public int IntSkipTestIdx = -1;
        //Offset_CP1_TD
        public int OffsetCp1TdIdx = -1;
        //Offset_CP1_BIST
        public int OffsetCp1BistIdx = -1;
        //Offset_CP1_FUNC
        public int OffsetCp1FuncIdx = -1;
        //Offset_CP2_TD
        public int OffsetCp2TdIdx = -1;
        //Offset_CP2_BIST
        public int OffsetCp2BistIdx = -1;
        //Offset_CP2_FUNC
        public int OffsetCp2FuncIdx = -1;
        //Offset_FT1_TD
        public int OffsetFt1TdIdx = -1;
        //Offset_FT1_BIST
        public int OffsetFt1BistIdx = -1;
        //Offset_FT1_FUNC
        public int OffsetFt1FuncIdx = -1;
        //Offset_FT2_TD
        public int OffsetFt2TdIdx = -1;
        //Offset_FT2_BIST
        public int OffsetFt2BistIdx = -1;
        //Offset_FT2_FUNC
        public int OffsetFt2FuncIdx = -1;
        //Offset_QA_TD
        public int OffsetQaTdIdx = -1;
        //Offset_QA_BIST
        public int OffsetQaBistIdx = -1;
        //Offset_QA_FUNC
        public int OffsetQaFuncIdx = -1;
        //Comment
        public int CommentIdx = -1;
        //BincutCalc_IDSrail
        public int BincutCalcIdSrailIdx = -1;
        public int BincutCalcIdSmax = -1;

        //Allow Equal
        public int AllowEqualIdx = -1;
        //Mono Delta
        public int MonoDeltaIdx = -1;
        public string BaseVoltage = "";
        public double StepSize;
        //VDD_CPU	VDD_GPU	VDD_SOC	VDD_CPU_SRAM	VDD_GPU_SRAM	VDD_FIXED	VDD_LOW	All Others	ATPG	MBIST	SPI/RTOS
        public List<Tuple<string, int>> TitleList = [];
        public List<BinningRow> Rows = [];

        public BinningTable() { }

        public BinningTable(BinningTable binningTable) : base(binningTable)
        {
            if (binningTable == null)
            {
                return;
            }

            Version = binningTable.Version;
            TitleRow1 = binningTable.TitleRow1;
            TitleRow2 = binningTable.TitleRow2;
            TitleRow3 = binningTable.TitleRow3;
            Job = binningTable.Job;
            BinnedIdx = binningTable.BinnedIdx;
            StartRowIdx = binningTable.StartRowIdx;
            DomainIdx = binningTable.DomainIdx;
            IdIdx = binningTable.IdIdx;
            ModeIdx = binningTable.ModeIdx;
            EqnBinIdx = binningTable.EqnBinIdx;
            BinXIdsMaxIdx = binningTable.BinXIdsMaxIdx;
            CpGbIdx = binningTable.CpGbIdx;
            Cp2GbIdx = binningTable.Cp2GbIdx;
            Ft1GbIdx = binningTable.Ft1GbIdx;
            Ft2GbIdx = binningTable.Ft2GbIdx;
            QaGbIdx = binningTable.QaGbIdx;
            HtolGbIdx = binningTable.HtolGbIdx;
            HtolGbRoomIdx = binningTable.HtolGbRoomIdx;
            HtolGbHotIdx = binningTable.HtolGbHotIdx;
            EqnIdx = binningTable.EqnIdx;
            ModeEqnIdx = binningTable.ModeEqnIdx;
            CIdx = binningTable.CIdx;
            MIdx = binningTable.MIdx;
            IdsMaxIdx = binningTable.IdsMaxIdx;
            FtMaxIdx = binningTable.FtMaxIdx;
            CpVMaxIdx = binningTable.CpVMaxIdx;
            CpVMinIdx = binningTable.CpVMinIdx;
            CpHvIdx = binningTable.CpHvIdx;
            FtHvIdx = binningTable.FtHvIdx;
            QaHvIdx = binningTable.QaHvIdx;
            SramthreshProductIdx = binningTable.SramthreshProductIdx;
            SramthreshCp1Idx = binningTable.SramthreshCp1Idx;
            SramthreshCp2Idx = binningTable.SramthreshCp2Idx;
            IntModeLIdx = binningTable.IntModeLIdx;
            IntModeHIdx = binningTable.IntModeHIdx;
            IntMfIdx = binningTable.IntMfIdx;
            IntOffsetIdx = binningTable.IntOffsetIdx;
            IntSkipTestIdx = binningTable.IntSkipTestIdx;
            OffsetCp1TdIdx = binningTable.OffsetCp1TdIdx;
            OffsetCp1BistIdx = binningTable.OffsetCp1BistIdx;
            OffsetCp1FuncIdx = binningTable.OffsetCp1FuncIdx;
            OffsetCp2TdIdx = binningTable.OffsetCp2TdIdx;
            OffsetCp2BistIdx = binningTable.OffsetCp2BistIdx;
            OffsetCp2FuncIdx = binningTable.OffsetCp2FuncIdx;
            OffsetFt1TdIdx = binningTable.OffsetFt1TdIdx;
            OffsetFt1BistIdx = binningTable.OffsetFt1BistIdx;
            OffsetFt1FuncIdx = binningTable.OffsetFt1FuncIdx;
            OffsetFt2TdIdx = binningTable.OffsetFt2TdIdx;
            OffsetFt2BistIdx = binningTable.OffsetFt2BistIdx;
            OffsetFt2FuncIdx = binningTable.OffsetFt2FuncIdx;
            OffsetQaTdIdx = binningTable.OffsetQaTdIdx;
            OffsetQaBistIdx = binningTable.OffsetQaBistIdx;
            OffsetQaFuncIdx = binningTable.OffsetQaFuncIdx;
            CommentIdx = binningTable.CommentIdx;
            BincutCalcIdSrailIdx = binningTable.BincutCalcIdSrailIdx;
            BincutCalcIdSmax = binningTable.BincutCalcIdSmax;
            AllowEqualIdx = binningTable.AllowEqualIdx;
            MonoDeltaIdx = binningTable.MonoDeltaIdx;
            BaseVoltage = binningTable.BaseVoltage;
            StepSize = binningTable.StepSize;
            TitleList = [.. binningTable.TitleList];
            Rows = [.. binningTable.Rows.Select(x => x.Copy())];
        }

        public BinningTable Copy()
        {
            return new BinningTable(this);
        }

        public VddBinDefComment GetAdjustPowerList()
        {
            var vddBinDefComment = new VddBinDefComment();
            var maxList = Rows.Where(x => x.RowData[CommentIdx].StartsWithIgnoreCase("Max PV")).Select(y => y.RowData[CommentIdx].ToUpper()).ToList();
            var minList = Rows.Where(x => x.RowData[CommentIdx].StartsWithIgnoreCase("Min PV")).Select(y => y.RowData[CommentIdx].ToUpper()).ToList();

            if (maxList.Count > 0)
            {
                vddBinDefComment.IsMax = true;
                maxList = [.. maxList.Distinct()];
                var modeList = new List<string>();
                foreach (string power in maxList)
                {
                    modeList.Add(_regex.Replace(power, "").Trim().TrimStart('(').TrimEnd(')').Replace('/', ','));
                }
                vddBinDefComment.MaxStr += string.Join("+", modeList);
            }

            if (minList.Count > 0)
            {
                minList = [.. minList.Distinct()];
                vddBinDefComment.IsMin = true;
                var modeList = new List<string>();
                foreach (string power in minList)
                {
                    modeList.Add(_regex2.Replace(power, "").Trim().TrimStart('(').TrimEnd(')').Replace('/', ','));
                }
                vddBinDefComment.MinStr += string.Join("+", modeList);
            }
            return vddBinDefComment;
        }

        public int GetColumnIndex(string express)
        {
            List<string> arr = [.. _regex3.Split(express)];
            if (arr.Any(x => x.EqualsIgnoreCase("DOMAIN")))
            {
                return DomainIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("MODE")))
            {
                return ModeIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("C")))
            {
                return CIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("CPGB")) ||
                arr.Any(x => x.EqualsIgnoreCase("BinningGB")))
            {
                return CpGbIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("CP2GB")))
            {
                return Cp2GbIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("FT_GB_ROOM")))
            {
                return Ft1GbIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("FT_GB_HOT")))
            {
                return Ft2GbIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("ATE_FQAGB")))
            {
                return QaGbIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("CPHV")))
            {
                return CpHvIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("FTHV")))
            {
                return FtHvIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("QAHV")))
            {
                return QaHvIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("INT_Mode_L")))
            {
                return IntModeLIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("INT_Mode_H")))
            {
                return IntModeHIdx;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("INT_MF")))
            {
                return IntMfIdx;
            }

            (bool flowControl, int value) = GetColumnIndexOffSet(arr);
            if (!flowControl)
            {
                return value;
            }

            if (arr.Any(x => x.EqualsIgnoreCase("CPVmax")) ||
                arr.Any(x => x.EqualsIgnoreCase("BinningVmax")))
            {
                return CpVMaxIdx;
            }

            return arr.Any(x => x.EqualsIgnoreCase("CPVmin")) ||
                arr.Any(x => x.EqualsIgnoreCase("BinningVmin"))
                ? CpVMinIdx
                : -1;
        }

        private (bool flowControl, int value) GetColumnIndexOffSet(List<string> arr)
        {
            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_CP1_TD")))
            {
                return (flowControl: false, value: OffsetCp1TdIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_CP1_BIST")))
            {
                return (flowControl: false, value: OffsetCp1BistIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_CP1_FUNC")))
            {
                return (flowControl: false, value: OffsetCp1FuncIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_CP2_TD")))
            {
                return (flowControl: false, value: OffsetCp2TdIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_CP2_BIST")))
            {
                return (flowControl: false, value: OffsetCp2BistIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_CP2_FUNC")))
            {
                return (flowControl: false, value: OffsetCp2FuncIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_FT1_TD")))
            {
                return (flowControl: false, value: OffsetFt1TdIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_FT1_BIST")))
            {
                return (flowControl: false, value: OffsetFt1BistIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_FT1_FUNC")))
            {
                return (flowControl: false, value: OffsetFt1FuncIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_FT2_TD")))
            {
                return (flowControl: false, value: OffsetFt2TdIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_FT2_BIST")))
            {
                return (flowControl: false, value: OffsetFt2BistIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_FT2_FUNC")))
            {
                return (flowControl: false, value: OffsetFt2FuncIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_QA_TD")))
            {
                return (flowControl: false, value: OffsetQaTdIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_QA_BIST")))
            {
                return (flowControl: false, value: OffsetQaBistIdx);
            }

            if (arr.Any(x => x.EqualsIgnoreCase("OFFSET_QA_FUNC")))
            {
                return (flowControl: false, value: OffsetQaFuncIdx);
            }

            return (flowControl: true, value: default);
        }

        public bool IsInterpolationSkip(string mode)
        {
            IEnumerable<BinningRow> rows = Rows.Where(x => x.RowData[ModeIdx].EqualsIgnoreCase(mode));
            foreach (BinningRow row in rows)
            {
                if (IntModeLIdx != -1 && IntModeHIdx != -1
                    && !string.IsNullOrEmpty(row.RowData[IntModeLIdx]) &&
                    !string.IsNullOrEmpty(row.RowData[IntModeHIdx]))
                {
                    string intSkipTest = row.RowData[IntSkipTestIdx];
                    return intSkipTest.EqualsIgnoreCase("True") ||
                           intSkipTest.EqualsIgnoreCase("Yes");
                }
            }
            return false;
        }

        public bool IsInterpolationSkip(string mode, out string domain, out string intModeL, out string intModeH, out string intSkipTest)
        {
            domain = "";
            intModeL = "";
            intModeH = "";
            intSkipTest = "";
            IEnumerable<BinningRow> rows = Rows.Where(x => x.RowData[ModeIdx].EqualsIgnoreCase(mode));
            foreach (BinningRow row in rows)
            {
                if (IntModeLIdx != -1 && IntModeHIdx != -1
                    && !string.IsNullOrEmpty(row.RowData[IntModeLIdx]) &&
                    !string.IsNullOrEmpty(row.RowData[IntModeHIdx]))
                {
                    domain = row.RowData[DomainIdx];
                    intModeL = row.RowData[IntModeLIdx];
                    intModeH = row.RowData[IntModeHIdx];
                    intSkipTest = row.RowData[IntSkipTestIdx];
                    if (!_regex4.IsMatch(intModeL) || !_regex4.IsMatch(intModeH))
                    {
                        return false;
                    }
                    return true;
                }
            }
            return false;
        }

        public List<string> GetPowerNames()
        {
            var powerNames = new List<string>();
            foreach (BinningRow row in Rows)
            {
                string domain = row.RowData[DomainIdx];
                string mode = row.RowData[ModeIdx];
                string powerName = "VDD_" + domain + "_" + mode;
                if (row.IsOtherRail)
                {
                    powerName = "VDD_" + domain;
                }

                if (!powerNames.Any(x => x.EqualsIgnoreCase(powerName)))
                {
                    powerNames.Add(powerName);
                }
            }
            return powerNames;
        }

        public Dictionary<string, string> GetModeVsPowerName()
        {
            var powerDic = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            foreach (BinningRow row in Rows)
            {
                string domain = row.RowData[DomainIdx];
                string mode = row.RowData[ModeIdx];
                string powerName = "VDD_" + domain;
                if (row.IsOtherRail)
                {
                    continue;
                }

                if (!powerDic.ContainsKey(mode.ToUpper()))
                {
                    powerDic.Add(mode.ToUpper(), powerName.ToUpper());
                }
            }
            return powerDic;
        }

        public List<string> GetModes()
        {
            var modes = new List<string>();
            foreach (BinningRow row in Rows)
            {
                if (!row.RowData[BinnedIdx].EqualsIgnoreCase("True"))
                {
                    continue;
                }

                string mode = row.RowData[ModeIdx];
                if (!string.IsNullOrEmpty(mode))
                {
                    modes.Add(mode);
                }
            }
            return [.. modes.Distinct()];
        }

        public List<PinInfo> GetModeVsPin()
        {
            var pinInfos = new List<PinInfo>();
            var groups = Rows.GroupBy(x => x.RowData[DomainIdx]).ToList();
            foreach (IGrouping<string, BinningRow> group in groups)
            {
                var modes = group.Select(x => x.RowData[ModeIdx]).Distinct().ToList();
                foreach (BinningRow row in group)
                {
                    string domain = row.RowData[DomainIdx];
                    string mode = row.RowData[ModeIdx];
                    string binned = row.RowData[BinnedIdx];
                    string powerName = "VDD_" + domain;
                    var pinInfo = new PinInfo
                    {
                        Mode = mode,
                        Pin = powerName,
                        Binned = binned
                    };
                    pinInfo.PinMode = modes.Count != 1 || binned.EqualsIgnoreCase("TRUE") ?
                        Combination.CombineByUnderLine(pinInfo.Pin, pinInfo.Mode) : pinInfo.Pin;
                    pinInfos.Add(pinInfo);
                }
            }
            var groups1 = pinInfos.GroupBy(x => x.PinMode).ToList();

            //foreach (var row in Rows)
            //{
            //    string domain = row.RowData[DomainIdx];
            //    string mode = row.RowData[ModeIdx];
            //    string binned = row.RowData[BinnedIdx];
            //    var powerName = "VDD_" + domain;
            //    var pinInfo = new PinInfo();
            //    pinInfo.Mode = mode;
            //    pinInfo.Pin = powerName;
            //    pinInfo.Binned = binned;
            //    pinInfo.PinMode = pinInfo.Binned.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase) ?
            //        Combination.CombineByUnderLine(pinInfo.Pin, pinInfo.Mode) : pinInfo.Pin;
            //    pinInfos.Add(pinInfo);
            //}

            //var groups1 = pinInfos.GroupBy(x => x.PinMode).ToList();
            //pinInfos = groups1.Select(x => x.First()).ToList();
            return groups1.ConvertAll(x => x.First());
        }

        #region check
        private void CheckAllowEqual()
        {
            if (AllowEqualIdx == -1 || ModeIdx == -1)
            {
                return;
            }

            var groups = Rows.GroupBy(x => x.RowData[ModeIdx]).ToList();
            for (int i = 0; i < groups.Count; i++)
            {
                IGrouping<string, BinningRow> group = groups[i];
                foreach (BinningRow row in group)
                {
                    string allowEqual = row.RowData[AllowEqualIdx];
                    if (!string.IsNullOrEmpty(allowEqual))
                    {
                        string domainSon = DomainIdx == -1 ? "" : row.RowData[DomainIdx];
                        IGrouping<string, BinningRow>? allowEqualMode = groups.Find(x => x.Key.EqualsIgnoreCase(allowEqual));
                        string? domainPar = null;
                        if (allowEqualMode?.Any() == true)
                        {
                            domainPar = DomainIdx == -1 ? "" : allowEqualMode.First().RowData[DomainIdx];
                        }

                        if ((!string.IsNullOrEmpty(domainPar) && !domainPar.EqualsIgnoreCase(domainSon)) ||
                            (CommentIdx != -1 && string.IsNullOrEmpty(row.RowData[CommentIdx])))
                        {
                            continue;
                        }

                        string previousMode = "";
                        if (i != 0)
                        {
                            previousMode = groups[i - 1].Key;
                        }

                        if (allowEqual != previousMode)
                        {
                            string errorMessage = $"The format of Allow Equal \" {allowEqual} \" was worng (Should be {previousMode})!!!";
                            AddError(BinCutErrorType.E_AllowEqual_01, SheetName, row.RowNum, AllowEqualIdx + 1, $"The format of Allow Equal \" {allowEqual} \" was worng (Should be {previousMode})!!!", [allowEqual, previousMode]);
                        }
                    }
                }
            }
        }

        private void CheckHtolgb()
        {
            for (int i = 0; i < Rows.Count; i++)
            {
                bool hasHtolRoGb = HtolGbIdx != -1;
                bool hasHtolGbRoom = HtolGbRoomIdx != -1;
                bool hasHtolGbHot = HtolGbHotIdx != -1;
                bool hasFtGbRoom = Ft1GbIdx != -1;
                bool hasFtGbHot = Ft2GbIdx != -1;

                double htolRoGb = GetValue(Rows[i].RowData, HtolGbIdx);
                double htolGbRoom = GetValue(Rows[i].RowData, HtolGbRoomIdx);
                double htolGbHot = GetValue(Rows[i].RowData, HtolGbHotIdx);
                double ftGbRoom = GetValue(Rows[i].RowData, Ft1GbIdx);
                double ftGbHot = GetValue(Rows[i].RowData, Ft2GbIdx);

                int rowIdx = Rows[i].RowNum;
                if (htolRoGb > ftGbRoom && hasHtolRoGb && hasFtGbRoom)
                {
                    //string errorMessage = $"The GB of HTOL {htolRoGb} is bigger than FT {ftGbRoom}";
                    AddError(BinCutErrorType.E_HTOLGb_01, SheetName, rowIdx, Ft1GbIdx + 1, $"The GB of HTOL {htolRoGb} is bigger than FT {ftGbRoom}", [htolRoGb.ToString(), ftGbRoom.ToString()]);
                }

                if (htolRoGb > ftGbHot && hasHtolRoGb && hasFtGbHot)
                {
                    //string errorMessage = $"The GB of HTOL {htolRoGb} is bigger than FT {ftGbHot}";
                    AddError(BinCutErrorType.W_HTOLGb_01, SheetName, rowIdx, Ft2GbIdx + 1, $"The GB of HTOL {htolRoGb} is bigger than FT {ftGbHot}", [htolRoGb.ToString(), ftGbHot.ToString()]);
                }

                if (htolGbRoom > ftGbRoom && hasHtolGbRoom && hasFtGbRoom)
                {
                    //string errorMessage = $"The GB of HTOL {htolGbRoom} is bigger than FT {ftGbRoom}";
                    AddError(BinCutErrorType.W_HTOLGb_02, SheetName, rowIdx, Ft1GbIdx + 1, $"The GB of HTOL {htolGbRoom} is bigger than FT {ftGbRoom}", [htolGbRoom.ToString(), ftGbRoom.ToString()]);
                }

                if (htolGbHot > ftGbRoom && hasHtolGbHot && hasFtGbHot)
                {
                    //string errorMessage = $"The GB of HTOL {htolGbHot} is bigger than FT {ftGbRoom}";
                    AddError(BinCutErrorType.W_HTOLGb_03, SheetName, rowIdx, Ft2GbIdx + 1, $"The GB of HTOL {htolGbHot} is bigger than FT {ftGbRoom}", [htolGbHot.ToString(), ftGbRoom.ToString()]);
                }
            }
        }

        private static double GetValue(IEnumerable<string> rowData, int htolGbIdx)
        {
            double value = 0;
            if (htolGbIdx == -1)
            {
                return value;
            }

            string text = rowData.ElementAt(htolGbIdx);
            _ = double.TryParse(text, out value);
            return value;
        }

        private void CheckSraMthresh()
        {
            //Check SRAMthresh_Product and SRAMthresh_binSearch in all Binning
            for (int i = 0; i < Rows.Count; i++)
            {
                int rowIdx = Rows[i].RowNum;
                if (Rows[i].IsOtherRail)
                {
                    continue;
                }

                if (SramthreshProductIdx != -1)
                {
                    if (Rows[i].RowData[SramthreshProductIdx].StartsWith('-') ||
                        Rows[i].RowData[SramthreshProductIdx] == "0")
                    {
                        //string errorMessage = "The SRAMthresh_Product value is not allowed <= 0 !!!";
                        AddError(BinCutErrorType.E_SRAMthresh_01, SheetName, rowIdx, SramthreshProductIdx + 1, "The SRAMthresh_Product value is not allowed <= 0 !!!");
                    }
                }

                if (SramthreshCp1Idx != -1)
                {
                    if (Rows[i].RowData[SramthreshCp1Idx].StartsWith('-') ||
                        Rows[i].RowData[SramthreshCp1Idx] == "0")
                    {
                        //string errorMessage = "The SRAMthresh_binSearch value is not allowed <= 0 !!!";
                        AddError(BinCutErrorType.E_SRAMthresh_02, SheetName, rowIdx, SramthreshCp1Idx + 1, "The SRAMthresh_binSearch (CP1) value is not allowed to be <= 0.");
                    }
                }

                if (SramthreshCp2Idx != -1)
                {
                    if (Rows[i].RowData[SramthreshCp2Idx].StartsWith('-') ||
                        Rows[i].RowData[SramthreshCp2Idx] == "0")
                    {
                        //string errorMessage = "The SRAMthresh_binSearch value is not allowed <= 0 !!!";
                        AddError(BinCutErrorType.E_SRAMthresh_03, SheetName, rowIdx, SramthreshCp2Idx + 1, "The SRAMthresh_binSearch (CP2) value is not allowed to be <= 0.");
                    }
                }
            }
        }

        private void CheckAllEqInterpolation()
        {
            if (ModeIdx == -1)
            {
                return;
            }

            IEnumerable<IGrouping<string, BinningRow>> rowGroups = Rows.GroupBy(x => x.RowData[ModeIdx]);
            foreach (IGrouping<string, BinningRow> group in rowGroups)
            {
                if (IntMfIdx != -1)
                {
                    if (group.Any(row => row.RowData.Count > IntMfIdx && !string.IsNullOrEmpty(row.RowData[IntMfIdx])))
                    {
                        CompareEachEquationData(group, IntMfIdx);
                    }
                }

                if (IntModeHIdx != -1)
                {
                    if (group.Any(row => row.RowData.Count > IntModeHIdx && !string.IsNullOrEmpty(row.RowData[IntModeHIdx])))
                    {
                        CompareEachEquationData(group, IntModeHIdx);
                    }
                }

                if (IntModeLIdx != -1)
                {
                    if (group.Any(row => row.RowData.Count > IntModeLIdx && !string.IsNullOrEmpty(row.RowData[IntModeLIdx])))
                    {
                        CompareEachEquationData(group, IntModeLIdx);
                    }
                }

                if (IntOffsetIdx != -1)
                {
                    if (group.Any(row => row.RowData.Count > IntOffsetIdx && !string.IsNullOrEmpty(row.RowData[IntOffsetIdx])))
                    {
                        CompareEachEquationData(group, IntOffsetIdx);
                    }
                }

                if (IntSkipTestIdx != -1)
                {
                    if (group.Any(row => row.RowData.Count > IntSkipTestIdx && !string.IsNullOrEmpty(row.RowData[IntSkipTestIdx])))
                    {
                        CompareEachEquationData(group, IntSkipTestIdx);
                    }
                }
            }
        }

        private void CheckAllEqSramThreashold()
        {
            if (ModeIdx == -1)
            {
                return;
            }

            IEnumerable<IGrouping<string, BinningRow>> rowGroups = Rows.GroupBy(x => x.RowData[ModeIdx]);
            foreach (IGrouping<string, BinningRow> group in rowGroups)
            {
                if (SramthreshCp1Idx != -1)
                {
                    if (group.Any(row => row.RowData.Count > SramthreshCp1Idx && !string.IsNullOrEmpty(row.RowData[SramthreshCp1Idx])))
                    {
                        CompareEachEquationData(group, SramthreshCp1Idx);
                    }
                }

                if (SramthreshCp2Idx != -1)
                {
                    if (group.Any(row => row.RowData.Count > SramthreshCp2Idx && !string.IsNullOrEmpty(row.RowData[SramthreshCp2Idx])))
                    {
                        CompareEachEquationData(group, SramthreshCp2Idx);
                    }
                }

                if (SramthreshProductIdx != -1)
                {
                    if (group.Any(row => row.RowData.Count > SramthreshProductIdx && !string.IsNullOrEmpty(row.RowData[SramthreshProductIdx])))
                    {
                        CompareEachEquationData(group, SramthreshProductIdx);
                    }
                }
            }
        }

        private void CompareEachEquationData(IGrouping<string, BinningRow> group, int idx)
        {
            for (int i = 0; i < group.Count(); i++)
            {
                string data1 = group.ElementAt(i).RowData[idx];
                if (!_regex4.IsMatch(data1) && (IntModeHIdx == idx || IntModeLIdx == idx))
                {
                    AddError(BinCutErrorType.E_Interpolation_03, SheetName, group.ElementAt(i).RowNum, idx + 1, $"The value of Int_Mode_L/Int_Mode_H {data1} is not a legal performance mode for {group.Key}!!!", [data1, group.Key]);
                }

                if (string.IsNullOrEmpty(data1))
                {
                    //string errorMessage = $"The value of performance {group.Key} is different with other equations!!!";
                    AddError(BinCutErrorType.E_Interpolation_01, SheetName, group.ElementAt(i).RowNum, idx + 1, $"The value of performance {group.Key} is different with other equations!!!", [group.Key]);
                    continue;
                }

                for (int j = i + 1; j < group.Count(); j++)
                {
                    string data2 = group.ElementAt(j).RowData[idx];
                    if (string.IsNullOrEmpty(data2))
                    {
                        continue;
                    }

                    if (!data1.EqualsIgnoreCase(data2))
                    {
                        //string errorMessage = $"The value of performance {group.Key} is different with other equations!!!";
                        AddError(BinCutErrorType.E_Interpolation_01, SheetName, group.ElementAt(i).RowNum, idx + 1, $"The value of performance {group.Key} is different with other equations!!!", [group.Key]);
                        break;
                    }
                }
            }
        }

        public void Check()
        {
            CheckAllowEqual();

            CheckHtolgb();

            CheckSraMthresh();

            CheckAllEqSramThreashold();

            CheckAllEqInterpolation();
        }
        #endregion

        public bool IsTheDomain(string parName, string sonName)
        {
            string parDomain = "";
            string sonDomain = "";
            foreach (BinningRow row in Rows)
            {
                string mode = row.RowData[ModeIdx];
                string domain = row.RowData[DomainIdx];
                if (mode.EqualsIgnoreCase(parName))
                {
                    parDomain = domain;
                }

                if (mode.EqualsIgnoreCase(sonName))
                {
                    sonDomain = domain;
                }

                if (!string.IsNullOrEmpty(parDomain) &&
                    !string.IsNullOrEmpty(sonDomain))
                {
                    return parDomain.EqualsIgnoreCase(sonDomain);
                }
            }
            return true;
        }

        public List<AllowEqualBase> GetAllowEqualList()
        {
            var allowEqualList = new List<AllowEqualBase>();
            for (int i = 0; i < Rows.Count - 1; i++)
            {
                var allowEqualBase = new AllowEqualBase
                {
                    Mode = Rows[i].RowData[ModeIdx].Trim().ToUpper(),
                    AllowEqual = Rows[i].RowData[AllowEqualIdx].Trim().ToUpper()
                };
                if (CommentIdx != -1)
                {
                    allowEqualBase.Comment = Rows[i].RowData[CommentIdx].Trim();
                }

                if (!string.IsNullOrEmpty(allowEqualBase.AllowEqual) || !string.IsNullOrEmpty(allowEqualBase.Comment))
                {
                    if (!allowEqualList.Exists(x => x.Mode.EqualsIgnoreCase(allowEqualBase.Mode)))
                    {
                        allowEqualList.Add(allowEqualBase);
                    }
                }
            }
            return allowEqualList;
        }

        public List<string> GetAllowEqualModes()
        {
            var allowEqDic = new Dictionary<string, int>();
            for (int i = 0; i < Rows.Count - 1; i++)
            {
                string allowEq = Rows[i].RowData[AllowEqualIdx].Trim().ToUpper();
                string intModeL = Rows[i].RowData[IntModeLIdx].Trim().ToUpper();
                string intModeH = Rows[i].RowData[IntModeHIdx].Trim().ToUpper();
                if (string.IsNullOrEmpty(allowEq) || !string.IsNullOrEmpty(intModeL) || !string.IsNullOrEmpty(intModeH))
                {
                    continue;
                }

                string mode = Rows[i].RowData[ModeIdx].Trim().ToUpper();
                //allowEqModes.Add(mode);
                if (allowEqDic.TryGetValue(mode, out int value))
                {
                    allowEqDic[mode] = ++value;
                }
                else
                {
                    allowEqDic.Add(mode, 1);
                }
            }
            return [.. allowEqDic.Where(x => x.Value > 1).Select(x => x.Key)];
        }

        public List<string> GetSingleEquationModes()
        {
            var modeDic = new Dictionary<string, int>();
            for (int i = 0; i < Rows.Count - 1; i++)
            {
                string mode = Rows[i].RowData[ModeIdx].Trim().ToUpper();
                string binned = Rows[i].RowData[BinnedIdx].Trim().ToUpper();
                if (binned != "TRUE")
                {
                    continue;
                }

                //allowEqModes.Add(mode);
                if (modeDic.TryGetValue(mode, out int value))
                {
                    modeDic[mode] = ++value;
                }
                else
                {
                    modeDic.Add(mode, 1);
                }
            }
            return [.. modeDic.Where(x => x.Value == 1).Select(x => x.Key)];
        }

        public List<Dictionary<string, bool>> GetInherits(List<string> powerNames)
        {
            var inherits = new List<Dictionary<string, bool>>();
            var oneInherit = new Dictionary<string, bool>();
            var domainRows = Rows.GroupBy(x => x.RowData[DomainIdx]).ToList();
            foreach (IGrouping<string, BinningRow> domainRow in domainRows)
            {
                var sepDomainRow = new List<List<BinningRow>>();
                var tmpDomainRow = new List<BinningRow>();
                foreach (BinningRow oneRow in domainRow)
                {
                    _ = int.TryParse(oneRow.RowData[IdIdx], out int id);
                    if ((id == 1 || id % 100000 == 1001) && tmpDomainRow.Count != 0)
                    {
                        sepDomainRow.Add(tmpDomainRow);
                        tmpDomainRow = [];
                    }
                    tmpDomainRow.Add(oneRow);
                }
                sepDomainRow.Add(tmpDomainRow);
                foreach (List<BinningRow> domainRowByMode in sepDomainRow)
                {
                    var modeRows = domainRowByMode.GroupBy(x => x.RowData[ModeIdx]).OrderBy(y => double.Parse(y.Last().RowData[IdIdx])).ToList();
                    IEnumerable<string> inheritModes = modeRows.Select(modeRow => modeRow.Last().RowData[ModeIdx].Trim().ToUpper())
                        .Where(curModeName => powerNames.Exists(x => x.Contains(curModeName)));
                    foreach (string inheritMode in inheritModes)
                    {
                        oneInherit.Add(inheritMode, false);
                    }
                    var tmpList = new Dictionary<string, bool>(oneInherit);
                    inherits.Add(tmpList);
                    oneInherit.Clear();
                }
            }
            return inherits;
        }

        public BinningRow? GetBinningRow(string mode)
        {
            foreach (BinningRow row in Rows)
            {
                if (row.RowData[ModeIdx].EqualsIgnoreCase(mode))
                {
                    return row;
                }
            }
            return null;
        }

        public List<List<string>> GetInheritList()
        {
            var inheritModeList = new List<List<string>>();
            List<string> modeList = Rows.ConvertAll(x => x.RowData[ModeIdx]);
            var distinctModeList = new List<string>();
            foreach (BinningRow row in Rows)
            {
                if (_regex4.IsMatch(row.RowData[ModeIdx]))
                {
                    string mode = _regex4.Match(row.RowData[ModeIdx]).Groups["pmode"].ToString();
                    string number = _regex4.Match(row.RowData[ModeIdx]).Groups["modenumber"].ToString();
                    if (!string.IsNullOrEmpty(mode) &&
                        !string.IsNullOrEmpty(number) &&
                        !distinctModeList.Contains(mode + number[0]))
                    {
                        distinctModeList.Add(mode + number[0]);
                    }
                }
            }
            foreach (string item in distinctModeList)
            {
                inheritModeList.Add([.. modeList.Where(x => x.StartsWithIgnoreCase(item)).Distinct()]);
            }

            return inheritModeList;
        }

        public int GetGbIdx(EnumJob enumJob)
        {
            int gbIdx = -1;
            switch ((int)enumJob)
            {
                case 0:
                    gbIdx = CpGbIdx;
                    break;
                case 1:
                    gbIdx = Cp2GbIdx;
                    break;
                case 2:
                    gbIdx = Ft1GbIdx;
                    break;
                case 3:
                    gbIdx = Ft2GbIdx;
                    break;
                case 4:
                    gbIdx = QaGbIdx;
                    break;
            }
            return gbIdx;
        }

        public int GetOtherIndex(string dictExpress)
        {
            int tempRowIdx = -1;
            for (int rowIdx = 0; rowIdx < Rows.Count; rowIdx++)
            {
                string modeInOthRai = Rows[rowIdx].RowData[ModeIdx];
                string domainInOthRai = Rows[rowIdx].RowData[DomainIdx];
                if (modeInOthRai.EqualsIgnoreCase(dictExpress))
                {
                    tempRowIdx = rowIdx;
                    break;
                }
                if (domainInOthRai.EqualsIgnoreCase(dictExpress))
                {
                    tempRowIdx = rowIdx;
                    break;
                }
            }
            return tempRowIdx;
        }
    }

    public class BinningRow : MyRow
    {
        public string Line = "";
        public List<string> RowData = [];
        public bool IsOtherRail;

        public BinningRow() { }

        public BinningRow(BinningRow binningRow) : base(binningRow)
        {
            if (binningRow == null)
            {
                return;
            }

            Line = binningRow.Line;
            RowData = [.. binningRow.RowData];
            IsOtherRail = binningRow.IsOtherRail;
        }

        public BinningRow Copy()
        {
            return new BinningRow(this);
        }

        public double GetIdsMax(BinningTable binningTable)
        {
            string idsMaxIdx = binningTable.IdsMaxIdx != -1 ? RowData[binningTable.IdsMaxIdx] : "";
            string binXIdsMax = binningTable.BinXIdsMaxIdx != -1 ? RowData[binningTable.BinXIdsMaxIdx] : "";
            double idsMaxValue;
            if (!string.IsNullOrEmpty(binXIdsMax))
            {
                _ = double.TryParse(binXIdsMax, out idsMaxValue);
            }
            else
            {
                _ = double.TryParse(idsMaxIdx, out idsMaxValue);
            }

            return idsMaxValue;
        }

        public string GetOtherRailDomainType(BinningTable binningTable, Dictionary<string, EnumPowerType> powerType, Dictionary<string, string> domainInOtherRail2Power)
        {
            //ex: CPUSRAM
            string domainName = RowData[binningTable.ModeIdx];
            if (powerType.ContainsKey(domainName.ToUpper()))
            {
            }
            else
            {
                domainName = RowData[binningTable.DomainIdx];
                if (powerType.ContainsKey(domainName.ToUpper()))
                {
                }
                else
                {
                    if (domainInOtherRail2Power.TryGetValue(domainName, out string? value))
                    {
                    }
                    else
                    {
                        if (powerType.ContainsKey("VDD_" + domainName.ToUpper()))
                        {
                        }
                        else
                        {
                            string errorMessage = string.Format("The domain : " + domainName + " in Other Rails can't be found.");
                            if (!ErrorReportManager.GetErrorList().Select(x => x.Message).Contains(errorMessage))
                            {
                                ErrorReportManager.AddError(BinCutErrorType.E_Business_02, "Datalog", 0, 0, string.Format("The domain : " + domainName + " in Other Rails can't be found."), [domainName]);
                            }
                        }
                    }
                }
            }
            return domainName;
        }
    }

    public class AllowEqualBase
    {
        public string Mode = "";
        public string AllowEqual = "";
        public string Comment = "";
    }

    public class PinInfo
    {
        public string Binned = "";
        public string Pin = "";
        public string Mode = "";
        public string PinMode = "";
    }
}
