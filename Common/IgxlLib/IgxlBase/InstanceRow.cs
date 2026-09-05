using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using CommonLib.Extension;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{TestName} ")]
    public class InstanceRow : IgxlRow
    {
        public string TestName { get; set; } = string.Empty;
        public string VbtType { get; set; } = string.Empty;
        public string VbtName { get; set; } = string.Empty;
        public string CalledAs { get; set; } = string.Empty;
        public string DcCategory { get; set; } = string.Empty;
        public string DcSelector { get; set; } = string.Empty;
        public string AcCategory { get; set; } = string.Empty;
        public string AcSelector { get; set; } = string.Empty;
        public string TimeSets { get; set; } = string.Empty;
        public string EdgeSets { get; set; } = string.Empty;
        public string PinLevels { get; set; } = string.Empty;
        public string MixedSignalTiming { get; set; } = string.Empty;
        public string Overlay { get; set; } = string.Empty;
        public string ArgList { get; set; } = string.Empty;
        public List<string> Args { get; set; } = [];
        public string Comment { get; set; } = string.Empty;
        public List<string> InitList { get; set; } = [];
        public List<string> PayloadList { get; set; } = [];
        public List<string> FinalJobs { get; set; } = [];

        public InstanceRow()
        {
            SheetName = "";
        }

        public InstanceRow(InstanceRow instanceRow) : base(instanceRow)
        {
            if (instanceRow == null)
            {
                return;
            }

            TestName = instanceRow.TestName;
            VbtType = instanceRow.VbtType;
            VbtName = instanceRow.VbtName;
            CalledAs = instanceRow.CalledAs;
            DcCategory = instanceRow.DcCategory;
            DcSelector = instanceRow.DcSelector;
            AcCategory = instanceRow.AcCategory;
            AcSelector = instanceRow.AcSelector;
            TimeSets = instanceRow.TimeSets;
            EdgeSets = instanceRow.EdgeSets;
            PinLevels = instanceRow.PinLevels;
            MixedSignalTiming = instanceRow.MixedSignalTiming;
            Overlay = instanceRow.Overlay;
            ArgList = instanceRow.ArgList;
            Comment = instanceRow.Comment;

            // Deep Copy Lists - Ensures complete isolation
            Args = instanceRow.Args != null ? [.. instanceRow.Args] : [];
            InitList = instanceRow.InitList != null ? [.. instanceRow.InitList] : [];
            PayloadList = instanceRow.PayloadList != null ? [.. instanceRow.PayloadList] : [];
            FinalJobs = instanceRow.FinalJobs != null ? [.. instanceRow.FinalJobs] : [];
        }

        public InstanceRow Copy()
        {
            return new InstanceRow(this);
        }

        public string GetArgument(string argument)
        {
            int index = ArgList.Split(',').ToList().FindIndex(s => s.EqualsIgnoreCase(argument));
            if (index != -1)
            {
                return Args[index];
            }
            return "";
        }

        public void SetArgument(string argument, string value)
        {
            int index = ArgList.Split(',').ToList().FindIndex(s => s.EqualsIgnoreCase(argument));
            if (index != -1)
            {
                if (index < Args.Count)
                {
                    Args[index] = value;
                }
                else
                {
                    Args.AddRange([.. Enumerable.Repeat("", index - Args.Count)]);
                    Args.Add(value);
                }
            }
        }

        public List<string> GetDifferences(InstanceRow instanceRow, bool includeSheetName = true)
        {
            var diffs = new List<string>();

            void Check(string label, string val1, string val2)
            {
                if (!Compare(val1, val2))
                {
                    diffs.Add(label);
                }
            }

            Check(instanceRow.AcCategory, AcCategory, instanceRow.AcCategory);
            Check(instanceRow.AcSelector, AcSelector, instanceRow.AcSelector);
            Check("ArgList", ArgList, instanceRow.ArgList);
            Check(instanceRow.CalledAs, CalledAs, instanceRow.CalledAs);
            Check(instanceRow.Comment, Comment, instanceRow.Comment);
            Check(instanceRow.DcCategory, DcCategory, instanceRow.DcCategory);
            Check(instanceRow.DcSelector, DcSelector, instanceRow.DcSelector);
            Check(instanceRow.EdgeSets, EdgeSets, instanceRow.EdgeSets);
            Check(instanceRow.MixedSignalTiming, MixedSignalTiming, instanceRow.MixedSignalTiming);
            Check(instanceRow.VbtName, VbtName, instanceRow.VbtName);
            Check(instanceRow.Overlay, Overlay, instanceRow.Overlay);
            Check(instanceRow.PinLevels, PinLevels, instanceRow.PinLevels);
            Check(instanceRow.TestName, TestName, instanceRow.TestName);
            Check(instanceRow.TimeSets, TimeSets, instanceRow.TimeSets);
            Check(instanceRow.VbtType, VbtType, instanceRow.VbtType);

            if (includeSheetName)
            {
                Check(instanceRow.SheetName!, SheetName!, instanceRow.SheetName!);
            }

            // List Comparisons
            if (!ListEquals(InitList, instanceRow.InitList))
            {
                diffs.Add("InitPat");
            }

            if (!ListEquals(PayloadList, instanceRow.PayloadList))
            {
                diffs.Add("PayloadPat");
            }

            // Args Comparison
            if (Args.Count != instanceRow.Args.Count)
            {
                diffs.Add("ArgsCount");
            }
            else
            {
                for (int i = 0; i < Args.Count; i++)
                {
                    if (string.IsNullOrEmpty(Args[i]) && string.IsNullOrEmpty(instanceRow.Args[i]))
                    {
                        continue;
                    }
                    if (!Args[i].EqualsIgnoreCase(instanceRow.Args[i]))
                    {
                        diffs.Add("ArgsValue");
                        break;
                    }
                }
            }

            return [.. diffs.Distinct()];
        }

        private static bool ListEquals(List<string> list1, List<string> list2)
        {
            if (list1.Count == 0 || list2.Count == 0)
            {
                return true;
            }

            return list1.Count == list2.Count && list1.SequenceEqual(list2, StringExtensions.IgnoreCase);
        }

        private static bool Compare(string row1, string row2)
        {
            return (string.IsNullOrEmpty(row1) && string.IsNullOrEmpty(row2)) ||
                   (!string.IsNullOrEmpty(row1) && !string.IsNullOrEmpty(row2) && row1.EqualsIgnoreCase(row2));
        }

        public List<string>? GetPinGrpFlags()
        {
            string pinGrp = GetArgument("Harv_FailFlag");
            if (string.IsNullOrEmpty(pinGrp) || pinGrp == "HarvestPinGrpFlagTable" || pinGrp == "HarvestPinFlag_Table")
            {
                return null;
            }

            var pinFlags = pinGrp.Split(';').Select(x => x.Split('(').Last().Trim(')')).ToList();
            return pinFlags;
        }

        public void MergeVbtMtdArgs(InstanceRow instanceRow)
        {
            var mtdArgsPlus = new List<string> { "Patset" };
            var mtdArgsSemicolon = new List<string> { "Calc_Eqn" };
            var mtdArgsSharp = new List<string>
            {
                "TestSequence",
                "CPUA_Flag_In_Pat",
                "DisableComparePins",
                "DisableConnectPins",
                "DisableFRC",
                "FRCPortName",
                "MeasV_PinS",
                "MeasF_PinS_SingleEnd",
                "MeasI_pinS",
                "MeasI_Range",
                "MeasI_WaitTime",
                "DigCap_Pin",
                "DigCap_DataWidth",
                "DigCap_Sample_Size",
                "DigSrc_Pin",
                "DigSrc_DataWidth",
                "DigSrc_Sample_Size",
                "DigSrc_Equation",
                "DigSrc_Assignment",
                "DigSrc_FlowForLoopIntegerName",
                "SpecialCalcValSetting",
                "InstSpecialSetting",
                "CUS_Str_MainProgram",
                "CUS_Str_DigCapData",
                "CUS_Str_DigSrcData",
                "MeasF_PinS_Differential",
                "ForceFunctional_Flag",
                "MeasF_WalkingStrobe_Flag",
                "MeasF_WalkingStrobe_StartV",
                "MeasF_WalkingStrobe_EndV",
                "MeasF_WalkingStrobe_StepVoltage",
                "MeasF_WalkingStrobe_BothVohVolDiffV",
                "MeasF_WalkingStrobe_interval",
                "MeasF_WalkingStrobe_miniFreq",
                "Meas_StoreName",
                "Interpose_PrePat",
                "Interpose_PreMeas",
                "Interpose_PostMeas",
                "Interpose_PostTest",
                "CharSetName",
                "ForceV_Val",
                "ForceI_Val",
                "UVI80_MeasV_WaitTime",
                "RAK_Flag",
                "WaitTime_VIRZ",
                "MSB_First_Flag",
                "BV_Enable",
                "DSSCSetup",
                "Validating_"
            };

            foreach (string arg in mtdArgsPlus)
            {
                if (string.IsNullOrEmpty(GetArgument(arg)) && string.IsNullOrEmpty(instanceRow.GetArgument(arg)))
                {
                    continue;
                }

                SetArgument(arg, instanceRow.GetArgument(arg) + "+" + GetArgument(arg));
            }
            foreach (string arg in mtdArgsSemicolon)
            {
                if (string.IsNullOrEmpty(GetArgument(arg)) && string.IsNullOrEmpty(instanceRow.GetArgument(arg)))
                {
                    continue;
                }

                SetArgument(arg, instanceRow.GetArgument(arg) + ";" + GetArgument(arg));
            }
            foreach (string arg in mtdArgsSharp)
            {
                if (string.IsNullOrEmpty(GetArgument(arg)) && string.IsNullOrEmpty(instanceRow.GetArgument(arg)))
                {
                    continue;
                }

                SetArgument(arg, instanceRow.GetArgument(arg) + "#" + GetArgument(arg));
            }

        }
    }
}
