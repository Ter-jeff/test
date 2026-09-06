using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.common.IgxlProgramMappingLib.Base;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan
{
    public class Characterization
    {
        #region property
        public string SheetName;
        public int RowNum;
        public Dictionary<string, int> ColumnDic = new Dictionary<string, int>();

        public List<int> ColNum(List<string> keys)
        {
            return (from key in keys where ColumnDic.ContainsKey(key) select ColumnDic[key]).ToList();
        }

        public List<int> ColNum(string key)
        {
            var result = new List<int>();
            if (key.ToLower() == "all")
            {
                result.AddRange(ColumnDic.Values.ToList());
            }

            if (ColumnDic.TryGetValue(key, out int value))
            {
                result.Add(value);
            }

            return result;
        }

        public bool OnlyHasPayload1
        {
            get
            {
                return string.IsNullOrEmpty(Payload2) &&
                       string.IsNullOrEmpty(Payload3) &&
                       string.IsNullOrEmpty(Payload4) &&
                       string.IsNullOrEmpty(Payload5) &&
                       string.IsNullOrEmpty(Init1) &&
                       string.IsNullOrEmpty(Init2) &&
                       string.IsNullOrEmpty(Init3) &&
                       string.IsNullOrEmpty(Init4) &&
                       string.IsNullOrEmpty(Init5) &&
                       string.IsNullOrEmpty(Init6) &&
                       string.IsNullOrEmpty(Init7) &&
                       string.IsNullOrEmpty(Init8) &&
                       string.IsNullOrEmpty(Init9) &&
                       string.IsNullOrEmpty(Init10);
            }
        }

        public string DcSelector { get; set; }
        public string Use { get; set; }

        public bool IsUse
        {
            get { return Use.ToLower().Trim() == "use"; }
        }

        public string Enableword { get; set; }
        public string Update { get; set; }
        public string Priority { get; set; }
        public string UserDef1 { get; set; }
        public string UserDef2 { get; set; }
        public string UserDef3 { get; set; }
        public string UserDef4 { get; set; }
        public string UserDef5 { get; set; }
        public string UserDef6 { get; set; }
        public string UserDef7 { get; set; }
        public string UserDef8 { get; set; }
        public string UserDef9 { get; set; }
        public string Group { get; set; }
        public string Category { get; set; }
        public string OriTpName { get; set; }
        public List<string> IsPatternNotExist { get; set; }
        public List<string> ExtraInits { get; set; }
        public List<string> ExtraPLs { get; set; }
        public string SimTpName
        {
            get
            {
                var nameList = new List<string> {
                    UserDef1,
                    UserDef2,
                    UserDef3,
                    Group,
                    Category,
                    UserDef4,
                    UserDef5,
                    UserDef6,
                    UserDef7,
                    UserDef8
                };
                if (UserDef9 != "")
                {
                    nameList.Add(UserDef9);
                }

                return string.Join("_", nameList);
            }
        }
        public string Init1 { get; set; }
        public string Init2 { get; set; }
        public string Init3 { get; set; }
        public string Init4 { get; set; }
        public string Init5 { get; set; }
        public string Init6 { get; set; }
        public string Init7 { get; set; }
        public string Init8 { get; set; }
        public string Init9 { get; set; }
        public string Init10 { get; set; }
        public string Payload1 { get; set; }
        public string Payload2 { get; set; }
        public string Payload3 { get; set; }
        public string Payload4 { get; set; }
        public string Payload5 { get; set; }

        public bool IsSelSram
        {
            get
            {
                const string regSelSram = "Selsr[a]*m";
                return Regex.IsMatch(UserDef9, regSelSram, RegexOptions.IgnoreCase);
            }
        }

        public Dictionary<string, string> InitPatterns
        {
            get
            {
                return new Dictionary<string, string>
                {
                    {"init1", Init1},
                    {"init2", Init2},
                    {"init3", Init3},
                    {"init4", Init4},
                    {"init5", Init5},
                    {"init6", Init6},
                    {"init7", Init7},
                    {"init8", Init8},
                    {"init9", Init9},
                    {"init10", Init10}
                };
            }
        }

        public Dictionary<string, string> PayloadPatterns
        {
            get
            {
                return new Dictionary<string, string>
                {
                    {"payload1", Payload1},
                    {"payload2", Payload2},
                    {"payload3", Payload3},
                    {"payload4", Payload4},
                    {"payload5", Payload5},
                };
            }
        }

        public Dictionary<string, string> AllPatterns  // todo: merge AllPatterns into AllNonEmptyPatterns
        {
            get
            {
                return new Dictionary<string, string>
                {
                    {"init1", Init1},
                    {"init2", Init2},
                    {"init3", Init3},
                    {"init4", Init4},
                    {"init5", Init5},
                    {"init6", Init6},
                    {"init7", Init7},
                    {"init8", Init8},
                    {"init9", Init9},
                    {"init10", Init10},
                    {"payload1", Payload1},
                    {"payload2", Payload2},
                    {"payload3", Payload3},
                    {"payload4", Payload4},
                    {"payload5", Payload5},
                };
            }
        }

        public List<Tuple<string, string>> AllNonEmptyPatterns
        {
            get
            {
                var tupleList = new List<Tuple<string, string>>();
                if (!string.IsNullOrEmpty(Init1))
                {
                    tupleList.Add(new Tuple<string, string>("init1", Init1));
                }

                if (!string.IsNullOrEmpty(Init2))
                {
                    tupleList.Add(new Tuple<string, string>("init2", Init2));
                }

                if (!string.IsNullOrEmpty(Init3))
                {
                    tupleList.Add(new Tuple<string, string>("init3", Init3));
                }

                if (!string.IsNullOrEmpty(Init4))
                {
                    tupleList.Add(new Tuple<string, string>("init4", Init4));
                }

                if (!string.IsNullOrEmpty(Init5))
                {
                    tupleList.Add(new Tuple<string, string>("init5", Init5));
                }

                if (!string.IsNullOrEmpty(Init6))
                {
                    tupleList.Add(new Tuple<string, string>("init6", Init6));
                }

                if (!string.IsNullOrEmpty(Init7))
                {
                    tupleList.Add(new Tuple<string, string>("init7", Init7));
                }

                if (!string.IsNullOrEmpty(Init8))
                {
                    tupleList.Add(new Tuple<string, string>("init8", Init8));
                }

                if (!string.IsNullOrEmpty(Init9))
                {
                    tupleList.Add(new Tuple<string, string>("init9", Init9));
                }

                if (!string.IsNullOrEmpty(Init10))
                {
                    tupleList.Add(new Tuple<string, string>("init10", Init10));
                }

                if (!string.IsNullOrEmpty(Payload1))
                {
                    tupleList.Add(new Tuple<string, string>("payload1", Payload1));
                }

                if (!string.IsNullOrEmpty(Payload2))
                {
                    tupleList.Add(new Tuple<string, string>("payload2", Payload2));
                }

                if (!string.IsNullOrEmpty(Payload3))
                {
                    tupleList.Add(new Tuple<string, string>("payload3", Payload3));
                }

                if (!string.IsNullOrEmpty(Payload4))
                {
                    tupleList.Add(new Tuple<string, string>("payload4", Payload4));
                }

                if (!string.IsNullOrEmpty(Payload5))
                {
                    tupleList.Add(new Tuple<string, string>("payload5", Payload5));
                }

                return tupleList;
            }
        }

        public bool IsFuncRow()
        {
            if (string.IsNullOrEmpty(UserDef1))
            {
                return false;
            }

            var funcUserDef1 = new List<string> { "HFL", "HFH", "HFLH", "DFTL", "DFTH", "DFTLH", "MCL", "MCH", "MCLH" };
            return funcUserDef1.Contains(UserDef1.ToUpper());
        }

        public bool IsHipRow()
        {
            if (string.IsNullOrEmpty(UserDef1))
            {
                return false;
            }

            var hipUserDef1 = new List<string> { "HAC", "HIO", "HFL", "HFH", "HFLH" };
            return hipUserDef1.Contains(UserDef1.ToUpper());
        }

        public Dictionary<string, string> NonEmptyPatterns
        {
            get
            {
                var nonEmptyPat = new Dictionary<string, string>();
                foreach (KeyValuePair<string, string> pat in AllPatterns.Where(pat => !string.IsNullOrEmpty(pat.Value)))
                {
                    nonEmptyPat[pat.Key] = pat.Value;
                }

                return nonEmptyPat;
            }
        }

        public bool HasAnyPayload
        {
            get
            {
                return NonEmptyPatterns.Any(pat =>
                {
                    string[] tokens = pat.Value.Split('_');
                    if (tokens.Length <= 4)
                    {
                        return false;
                    }

                    return Regex.IsMatch(tokens[3], "^FU|PL", RegexOptions.IgnoreCase);
                });
            }
        }

        public virtual string TpName { get; set; }
        public virtual string IpUse1 { get; set; }
        public virtual string IpUse2 { get; set; }
        public virtual string IpUse3 { get; set; }
        public virtual string IpUse4 { get; set; }
        public virtual string IpUse5 { get; set; }
        public string OtherSupplies { get; set; }
        public string XStepSize { get; set; }
        public string YStepSize { get; set; }
        public string YAxis { get; set; }

        public string ApplyVoltageFromBinCut { get; set; }

        public List<ShmooSpec> PrimaryShmooXList
        {
            // primary shmoo is the first shmoo who is sweeping
            get
            {
                var primaryShmooList = new List<ShmooSpec>();

                ShmooSpec primaryShmoo = PowerSupplyX.FirstOrDefault(shmoo => shmoo.IsSweep);

                if (primaryShmoo == null)
                {
                    return primaryShmooList;
                }

                primaryShmooList.Add(primaryShmoo);

                if (primaryShmoo.Step == "")
                {
                    primaryShmoo.Step = "0.005";
                }

                primaryShmooList.AddRange(
                    PowerSupplyX.Where(spec => spec.Name != primaryShmoo.Name)
                        .Where(spec => spec.IsSweep)
                        .Where(spec => spec.IsPowerPin || spec.Type == "VIH" || spec.Type == "VIL")
                        .Where(spec => primaryShmoo.Start == spec.Start && primaryShmoo.Stop == spec.Stop)
                    );
                return primaryShmooList;
            }
        }

        public List<ShmooSpec> PrimaryShmooYList
        {
            // primary shmoo is the first shmoo who is sweeping
            get
            {
                var primaryShmooList = new List<ShmooSpec>();

                ShmooSpec primaryShmoo = PowerSupplyY.FirstOrDefault(shmoo => shmoo.IsSweep);

                if (primaryShmoo == null)
                {
                    return primaryShmooList;
                }

                primaryShmooList.Add(primaryShmoo);

                if (primaryShmoo.Step == "")
                {
                    primaryShmoo.Step = "0.005";
                }

                primaryShmooList.AddRange(
                    PowerSupplyY.Where(spec => spec.Name != primaryShmoo.Name)
                        .Where(spec => spec.IsSweep)
                        .Where(spec => spec.IsPowerPin || spec.Type == "VIH" || spec.Type == "VIL")
                        .Where(spec => primaryShmoo.Start == spec.Start && primaryShmoo.Stop == spec.Stop)
                    );
                return primaryShmooList;
            }
        }

        public List<ShmooSpec> PrimaryShmooZList
        {
            // primary shmoo is the first shmoo who is sweeping
            get
            {
                var primaryShmooList = new List<ShmooSpec>();

                ShmooSpec primaryShmoo = PowerSupplyZ.FirstOrDefault(shmoo => shmoo.IsSweep);

                if (primaryShmoo == null)
                {
                    return primaryShmooList;
                }

                primaryShmooList.Add(primaryShmoo);

                if (primaryShmoo.Step == "")
                {
                    primaryShmoo.Step = "0.005";
                }

                primaryShmooList.AddRange(
                    PowerSupplyZ.Where(spec => spec.Name != primaryShmoo.Name)
                        .Where(spec => spec.IsSweep)
                        .Where(spec => spec.IsPowerPin || spec.Type == "VIH" || spec.Type == "VIL")
                        .Where(spec => primaryShmoo.Start == spec.Start && primaryShmoo.Stop == spec.Stop)
                    );
                return primaryShmooList;
            }
        }

        public List<ShmooSpec> PowerSupplyX { get; set; }

        public List<ShmooSpec> PowerSupplyY { get; set; }

        public List<ShmooSpec> PowerSupplyZ { get; set; }

        public List<int> GetShmooPinIndex(string name)
        {
            var indexList = new List<int>();
            ShmooSpec index = PowerSupplyX.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.CurrentCultureIgnoreCase));
            if (index != null)
            {
                indexList.Add(index.ColIndex);
            }

            return indexList;
        }

        public List<ShmooSpec> TrackingSpecX
        {
            // tracking specs are power supply who ...
            //    1) not the primary shmoo 
            //    2) is sweep
            //    3) has different start / stop
            get
            {
                ShmooSpec primaryShmoo = PrimaryShmooXList.FirstOrDefault();
                if (primaryShmoo == null)
                {
                    return new List<ShmooSpec>();
                }

                var trackingSpecs = new List<ShmooSpec>();

                foreach (ShmooSpec spec in PowerSupplyX
                    .Where(spec => spec.Name != primaryShmoo.Name)
                    .Where(spec => spec.IsSweep)
                    .Where(spec => spec.IsPowerPin || spec.Type == "VIH" || spec.Type == "VIL")
                    .Where(spec => primaryShmoo.Start != spec.Start || primaryShmoo.Stop != spec.Stop)
                    .Where(spec => spec.Start != spec.Stop))
                {
                    if (spec.Step != "" && Math.Abs(double.Parse(spec.Step)) < 1E-10)
                    {
                        continue;
                    }

                    trackingSpecs.Add(spec);

                    if (spec.Step != "")
                    {
                        continue;
                    }

                    const string outString = "Step is not empty besides primary shmoo!";
                    ErrorManager.AddError(ErrorType.ErrorShmooRange, SheetName, RowNum, 1, Use, outString);
                    ErrorReportManager.AddError(CharErrorType.E_ErrorShmooRange_01, SheetName, RowNum, 1, []);
                }

                return trackingSpecs;
            }
        }

        public List<ShmooSpec> TrackingSpecY
        {
            get
            {
                ShmooSpec primaryShmoo = PrimaryShmooYList.FirstOrDefault();
                if (primaryShmoo == null)
                {
                    return new List<ShmooSpec>();
                }

                var trackingSpecs = new List<ShmooSpec>();

                foreach (ShmooSpec spec in PowerSupplyY
                    .Where(spec => spec.Name != primaryShmoo.Name)
                    .Where(spec => spec.IsSweep)
                    .Where(spec => spec.IsPowerPin || spec.Type == "VIH" || spec.Type == "VIL")
                    .Where(spec => primaryShmoo.Start != spec.Start || primaryShmoo.Stop != spec.Stop))
                {
                    trackingSpecs.Add(spec);

                    if (spec.Step != "")
                    {
                        continue;
                    }

                    const string outString = "Step is not empty besides primary shmoo! ";
                    ErrorManager.AddError(ErrorType.ErrorShmooRange, SheetName, RowNum, 1, Use, outString);
                    ErrorReportManager.AddError(CharErrorType.E_ErrorShmooRange_01, SheetName, RowNum, 1, []);
                }

                return trackingSpecs;
            }
        }

        public List<ShmooSpec> TrackingSpecZ
        {
            get
            {
                ShmooSpec primaryShmoo = PrimaryShmooYList.FirstOrDefault();
                if (primaryShmoo == null)
                {
                    return new List<ShmooSpec>();
                }

                var trackingSpecs = new List<ShmooSpec>();

                foreach (ShmooSpec spec in PowerSupplyZ
                    .Where(spec => spec.Name != primaryShmoo.Name)
                    .Where(spec => spec.IsSweep)
                    .Where(spec => spec.IsPowerPin || spec.Type == "VIH" || spec.Type == "VIL")
                    .Where(spec => primaryShmoo.Start != spec.Start || primaryShmoo.Stop != spec.Stop))
                {
                    trackingSpecs.Add(spec);

                    if (spec.Step != "")
                    {
                        continue;
                    }

                    const string outString = "Step is not empty besides primary shmoo! ";
                    ErrorManager.AddError(ErrorType.ErrorShmooRange, SheetName, RowNum, 1, Use, outString);
                    ErrorReportManager.AddError(CharErrorType.E_ErrorShmooRange_01, SheetName, RowNum, 1, []);
                }

                return trackingSpecs;
            }
        }

        public List<Pin> Pins { get; set; }
        public string RefFreq { get; set; }
        public string ShiftFreq { get; set; }
        public string Retention { get; set; }
        public string TestCondition { get; set; }
        public string Search { get; set; }
        public string Nominal { get; set; }
        public string Htol { get; set; }
        public string Ttr { get; set; }
        public string PowerRunScenario { get; set; }
        public string DcCateName;
        public string AcStepName;

        public int MeasSeq
        {
            get
            {
                if (Regex.Matches(TpName, @"_M(?<MeasSeq>\d+)_*", RegexOptions.IgnoreCase).Count == 0)
                {
                    return 999;
                }

                string measseqStr = Regex.Matches(TpName, @"_M(?<MeasSeq>\d+)_*", RegexOptions.IgnoreCase)[0].Groups["MeasSeq"].ToString();
                return int.Parse(measseqStr);
            }
        }

        public string PinString { get; set; }  // is a replica of USERDEF4
        public string PinName { get; set; }
        public int DiffType { get; set; }
        public bool IsUseRtosCmd { get; set; }
        public string SuspendDatalog { get; set; }

        public string TimeSet { get; set; }
        public string HarvFstp { get; set; }
        public string SiteFlag { get; set; }
        public string FailFlag { get; set; }
        public string FailInfo { get; set; }
        public string Dfc { get; set; }
        public string Environment { get; set; }
        public string Burst { get; set; }
        public bool DigSrcShiftOrder { get; set; }
        public string BypassShmooHole { get; set; }
        public MappingKey MappingKey { get; set; }
        public MappingSpec MappingSpec { get; set; }
        public List<PatternCell> PatternCellList { get; set; }
        public string DigSrc { get; set; }
        public string RetentionRamp { get; set; }
        public string AdaptiveCooling { get; set; }
        public string StageCp1 { get; set; }
        public string StageCp2 { get; set; }
        public string StageFt1 { get; set; }
        public string StageFt2 { get; set; }
        public string Die { get; set; }
        public string DigSrcBitSize
        {
            get
            {
                return string.Join(",", PatternCellList
                                        .Where(x => !string.IsNullOrEmpty(x.DigSrcBitSize))
                                        .Select(x => $"{x.Name}:{x.DigSrcBitSize}"));
            }
        }
        public string DigSrcSeg
        {
            get
            {
                return string.Join(",", PatternCellList
                                        .Where(x => !string.IsNullOrEmpty(x.DigSrcSeg))
                                        .Select(x => $"{x.Name}:{x.DigSrcSeg}"));
            }
        }
        public string DigSrcPin
        {
            get
            {
                return string.Join(",", PatternCellList
                                        .Where(x => !string.IsNullOrEmpty(x.DigSrcPin))
                                        .Select(x => $"{x.Name}:{x.DigSrcPin}"));
            }
        }
        public string DigSrcEq
        {
            get
            {
                return string.Join(",", PatternCellList
                                        .Where(x => !string.IsNullOrEmpty(x.DigSrcEq))
                                        .Select(x => $"{x.Name}:{x.DigSrcEq}"));
            }
        }
        public string OneTimeInit { get; set; }
        public string FreeRunningClock { get; set; }
        public string UserFunction { get; set; }
        public string HarvestPinGrpOtherFail { get; set; }
        public string EnableCoreHarvest { get; set; }
        public string EnableCoreMask { get; set; }
        public string PinGrpSpecifyMask { get; set; }
        public string SsnSpecifyMask { get; set; }

        public string LevelsByBlock
        {
            get
            {
                string payloadBlock = !string.IsNullOrEmpty(Payload1) && Payload1.Split('_').Length > 4 ? Payload1.Split('_')[4] : "HardIP";
                if (payloadBlock.StartsWith("TD", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Scan";
                }

                if (payloadBlock.StartsWith("SA", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Scan";
                }

                if (payloadBlock.StartsWith("SC", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Scan";
                }

                if (payloadBlock.StartsWith("CH", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Scan";
                }

                if (payloadBlock.StartsWith("B", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Mbist";
                }
                return payloadBlock;
            }
        }
        #endregion
        #region constructor

        public Characterization()
        {
            PowerSupplyZ = new List<ShmooSpec>();
            PowerSupplyY = new List<ShmooSpec>();
            PowerSupplyX = new List<ShmooSpec>();
            Pins = new List<Pin>();
            ExtraInits = new List<string>();
            ExtraPLs = new List<string>();
            IsPatternNotExist = new List<string>();
        }

        public Characterization(Characterization charRow)
        {
            Copy(charRow);
        }

        #endregion

        #region methods

        public void Copy(Characterization item)
        {
            SheetName = item.SheetName;
            RowNum = item.RowNum;
            DcSelector = item.DcSelector;
            Use = item.Use;
            Enableword = item.Enableword;
            Update = item.Update;
            Priority = item.Priority;
            UserDef1 = item.UserDef1;
            UserDef2 = item.UserDef2;
            UserDef3 = item.UserDef3;
            UserDef4 = item.UserDef4;
            UserDef5 = item.UserDef5;
            UserDef6 = item.UserDef6;
            UserDef7 = item.UserDef7;
            UserDef8 = item.UserDef8;
            UserDef9 = item.UserDef9;
            Group = item.Group;
            Category = item.Category;
            OriTpName = item.OriTpName;
            TpName = item.TpName;
            Init1 = item.Init1;
            Init2 = item.Init2;
            Init3 = item.Init3;
            Init4 = item.Init4;
            Init5 = item.Init5;
            Init6 = item.Init6;
            Init7 = item.Init7;
            Init8 = item.Init8;
            Init9 = item.Init9;
            Init10 = item.Init10;
            Payload1 = item.Payload1;
            Payload2 = item.Payload2;
            Payload3 = item.Payload3;
            Payload4 = item.Payload4;
            Payload5 = item.Payload5;
            IpUse1 = item.IpUse1;
            IpUse2 = item.IpUse2;
            IpUse3 = item.IpUse3;
            IpUse4 = item.IpUse4;
            IpUse5 = item.IpUse5;
            OtherSupplies = item.OtherSupplies;
            XStepSize = item.XStepSize;
            YStepSize = item.YStepSize;
            YAxis = item.YAxis;
            Htol = item.Htol;
            Ttr = item.Ttr;
            PowerRunScenario = item.PowerRunScenario;
            RefFreq = item.RefFreq;
            ShiftFreq = item.ShiftFreq;
            Retention = item.Retention;
            TimeSet = item.TimeSet;
            HarvFstp = item.HarvFstp;
            SiteFlag = item.SiteFlag;
            FailFlag = item.FailFlag;
            FailInfo = item.FailInfo;
            Environment = item.Environment;
            Burst = item.Burst;
            DigSrcShiftOrder = item.DigSrcShiftOrder;
            BypassShmooHole = item.BypassShmooHole;
            TestCondition = item.TestCondition;
            Search = item.Search;
            Nominal = item.Nominal;
            DcCateName = item.DcCateName;
            AcStepName = item.AcStepName;
            PinName = item.PinName;
            Dfc = item.Dfc;
            ApplyVoltageFromBinCut = item.ApplyVoltageFromBinCut;
            IsPatternNotExist = new List<string>();
            foreach (string issue in item.IsPatternNotExist)
            {
                IsPatternNotExist.Add(issue);
            }
            IsUseRtosCmd = item.IsUseRtosCmd;
            SuspendDatalog = item.SuspendDatalog;
            ColumnDic = new Dictionary<string, int>();
            foreach (KeyValuePair<string, int> set in item.ColumnDic)
            {
                ColumnDic.Add(set.Key, set.Value);
            }

            PowerSupplyX = new List<ShmooSpec>();
            foreach (ShmooSpec powerX in item.PowerSupplyX)
            {
                var newitem = new ShmooSpec(powerX);
                PowerSupplyX.Add(newitem);
            }

            PowerSupplyY = new List<ShmooSpec>();
            foreach (ShmooSpec powerY in item.PowerSupplyY)
            {
                var newitem = new ShmooSpec(powerY);
                PowerSupplyY.Add(newitem);
            }

            PowerSupplyZ = new List<ShmooSpec>();
            foreach (ShmooSpec powerZ in item.PowerSupplyZ)
            {
                var newitem = new ShmooSpec(powerZ);
                PowerSupplyZ.Add(newitem);
            }

            Pins = new List<Pin>();
            Pins.AddRange(item.Pins);

            ExtraInits = new List<string>();
            ExtraInits.AddRange(item.ExtraInits);

            ExtraPLs = new List<string>();
            ExtraPLs.AddRange(item.ExtraPLs);
            MappingKey = item.MappingKey;
            MappingSpec = item.MappingSpec;
            DigSrc = item.DigSrc;
            PatternCellList = item.PatternCellList;
            RetentionRamp = item.RetentionRamp;
            AdaptiveCooling = item.AdaptiveCooling;
            OneTimeInit = item.OneTimeInit;
            FreeRunningClock = item.FreeRunningClock;
            //PatternCellList = new List<PatternCell>();
            //foreach (var PatternCell in item.PatternCellList)
            //{
            //    var newitem = new PatternCell(PatternCell);
            //    PatternCellList.Add(newitem);
            //}
            UserFunction = item.UserFunction;
            HarvestPinGrpOtherFail = item.HarvestPinGrpOtherFail;
            EnableCoreHarvest = item.EnableCoreHarvest;
            EnableCoreMask = item.EnableCoreMask;
            PinGrpSpecifyMask = item.PinGrpSpecifyMask;
            SsnSpecifyMask = item.SsnSpecifyMask;
            StageCp1 = item.StageCp1;
            StageCp2 = item.StageCp2;
            StageFt1 = item.StageFt1;
            StageFt2 = item.StageFt2;
            Die = item.Die;
        }

        public void SpliteSgmt32()
        {
            UserDef6 += "b";
            TpName = SimTpName;
        }

        public string GetLevelLabel(string axis)
        {
            //if (IsSelSram)
            if (!IsHipRow())
            {
                return "GlobalSpec";
            }

            string levelLabel = "";

            switch (axis.ToUpper())
            {
                case "X":
                    levelLabel = PrimaryShmooXList.Exists(p => p.IsValt) ? "Valt" : "Vmain";
                    break;
                case "Y":
                    levelLabel = PrimaryShmooYList.Exists(p => p.IsValt) ? "Valt" : "Vmain";
                    break;
                case "Z":
                    levelLabel = PrimaryShmooZList.Exists(p => p.IsValt) ? "Valt" : "Vmain";
                    break;
            }
            return levelLabel;
        }

        #endregion
    }
}
