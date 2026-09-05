using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using CommonLib.Extension;

using CommonReaderLib;
using CommonReaderLib.PatternListCsv;

using LogLib.Utility;

namespace DebugPlanReaderLib.DebugPlan
{
    public class AiTestPlanRow : MyRow
    {
        #region Constructor
        public AiTestPlanRow(string sourceSheetName)
        {
            SheetName = sourceSheetName;
            Inits = new List<PatternDate>();
            Payloads = new List<PatternDate>();
            Pins = new List<Pin>();
        }
        #endregion

        public EnumDataLoggingSettingType EnumDataLoggingSettingType
        {
            get
            {
                if (DataLoggingSetting.Contains("DFC"))
                {
                    return EnumDataLoggingSettingType.Dfc;
                }

                if (DataLoggingSetting.Contains("FC"))
                {
                    return EnumDataLoggingSettingType.Fc;
                }

                return EnumDataLoggingSettingType.Na;
            }
        }


        public List<string> GetTimeSetsByPayloads(PatternListSheet patternListSheet)
        {
            var timeSets = new List<string>();
            foreach (PatternDate payload in Payloads)
            {
                if (patternListSheet.Rows.Exists(x =>
                        x.Pattern.Equals(payload.OriName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    string timeSet = Path.GetFileNameWithoutExtension(patternListSheet.Rows
                        .Find(x => x.Pattern.Equals(payload.OriName, StringComparison.CurrentCultureIgnoreCase))
                        .TimeSet);
                    timeSets.Add(timeSet);
                }
            }

            return timeSets;
        }

        #region Property

        public string UseNotUse { set; get; } = "";
        public string Comment { set; get; } = "";
        public string TestInstanceName { set; get; } = "";
        public string AiType { set; get; } = "";
        public string DataLoggingSetting { set; get; } = "";
        public string Timeset { set; get; } = "";
        public string VoltageCategory { set; get; } = "";
        public string Order { set; get; } = "";
        public string Search { set; get; } = "";
        public string TempCondition { set; get; } = "";
        public string FailCycleEachPoint { set; get; } = "";
        public List<PatternDate> Inits { get; set; } = new List<PatternDate>();
        public List<PatternDate> Payloads { get; set; } = new List<PatternDate>();
        public List<PatternDate> Patterns
        {
            get
            {
                List<PatternDate> result = new List<PatternDate>();
                result.AddRange(Inits);
                result.AddRange(Payloads);
                return result;
            }
        }
        public string MappingPattern
        {
            get
            {
                if (Payloads.Any())
                {
                    return Payloads[0].OriName;
                }
                else if (Patterns.Any())
                {
                    return Patterns.Last().OriName;
                }
                return "";
            }
        }
        public int IndexMappingPattern { get; set; } = -1;
        public List<Pin> Pins { get; set; } = new List<Pin>();
        public string PatSetName { get; set; } = "";
        public string USL { get; set; } = "";
        public string LSL { get; set; } = "";
        public string USL_LSL
        {
            get
            {
                if (!string.IsNullOrEmpty(USL) && !string.IsNullOrEmpty(LSL))
                {
                    double uslValue;
                    double lslValue;
                    string usl = USL.ConvertNumber();
                    string lsl = LSL.ConvertNumber();

                    if (double.TryParse(usl, out uslValue) && double.TryParse(lsl, out lslValue))
                    {
                        // USL:0.7535; LSL: 0.685
                        return string.Format("USL:{0};LSL:{1}", uslValue, lslValue);
                    }
                }
                return "";
            }
        }
        public string AcCategory { get; set; } = "";
        public string PowerRunScenario { get; set; } = "";
        public string Retention { get; set; } = "";
        public List<string> RetentionList
        {
            get { return Retention.Split(',').Select(x => x.Trim()).ToList(); }
        }
        public string DigSrc { get; set; } = "";
        public string DigSrcBitSize { get; set; } = "";
        public string DigSrcPin { get; set; } = "";
        public string DigSrcEQ { get; set; } = "";
        public string DigSrcSeg { get; set; } = "";

        public string DcCategory
        {
            get
            {
                List<string> arr = VoltageCategory.Split(' ', '_').ToList();
                string last = arr.Last();
                if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase) ||
                    last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase) ||
                    last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return string.Join("_", arr.GetRange(0, arr.Count - 1));
                }

                return VoltageCategory;
            }
        }

        public string GetVoltageCategory()
        {
            string last = VoltageCategory.Split(' ', '_').Last();
            if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase))
            {
                return "HV";
            }

            if (last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase))
            {
                return "LV";
            }

            if (last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
            {
                return "NV";
            }

            return "NV";
        }

        public string DcSelector
        {
            get
            {
                string last = VoltageCategory.Split(' ', '_').Last();
                if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Max";
                }

                if (last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Min";
                }

                if (last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Typ";
                }

                return "Typ";
            }
        }

        public string TestName
        {
            get
            {
                //var last = VoltageCategory.Split(' ', '_').Last();
                //var dcSelector = "NV";
                //if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase))
                //{
                //    dcSelector = "HV";
                //}

                //if (last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase))
                //{
                //    dcSelector = "LV";
                //}

                //if (last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
                //{
                //    dcSelector = "NV";
                //}

                //return PatSetName + "_" + dcSelector;
                return PatSetName;
            }
        }

        public EnumAiType EnumAiType
        {
            get
            {
                if (AiType.EndsWith("Data log", StringComparison.CurrentCultureIgnoreCase))
                {
                    return EnumAiType.Datalog;
                }

                if (AiType.Equals("1D", StringComparison.CurrentCultureIgnoreCase))
                {
                    return EnumAiType.Shmoo1D;
                }

                if (AiType.Equals("2D", StringComparison.CurrentCultureIgnoreCase))
                {
                    return EnumAiType.Shmoo2D;
                }

                return EnumAiType.Datalog;
            }
        }

        public string Parameter
        {
            get
            {
                if (EnumAiType == EnumAiType.Datalog)
                {
                    return GetTestName();
                }

                if (EnumAiType == EnumAiType.Shmoo1D)
                {
                    return GetTestName() + " " + CharName;
                }

                if (EnumAiType == EnumAiType.Shmoo2D)
                {
                    return GetTestName() + " " + CharName;
                }

                return GetTestName();
            }
        }

        public string CharName
        {
            get
            {
                if (EnumAiType == EnumAiType.Datalog)
                {
                    return "";
                }

                if (EnumAiType == EnumAiType.Shmoo1D)
                {
                    return "Char_1D_" + GetTestName();
                }

                if (EnumAiType == EnumAiType.Shmoo2D)
                {
                    return "Char_2D_" + GetTestName();
                }

                return "";
            }
        }

        public string SelsramDssc { get; set; } = "";
        public List<string> AcCategoryMapping { get; set; } = new List<string>();
        public List<string> TimesetMapping { get; set; } = new List<string>();
        public List<string> DcLevelsMapping { get; set; } = new List<string>();

        internal string GetBlock()
        {
            string firstSeg = !string.IsNullOrEmpty(MappingPattern) ? MappingPattern.Split('_')[4] : "HardIP";
            //var firstSeg = DcCategory.Split('_').First();
            if (firstSeg.StartsWith("TD", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Scan";
            }

            if (firstSeg.StartsWith("SA", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Scan";
            }

            if (firstSeg.StartsWith("SC", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Scan";
            }

            if (firstSeg.StartsWith("CH", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Scan";
            }

            if (firstSeg.StartsWith("B", StringComparison.CurrentCultureIgnoreCase))
            {
                return "Mbist";
            }
            if (MappingPattern.Split('_').Length >= 11)
            {
                LogHelper.Debug($"MappingPattern: {MappingPattern}");
                LogHelper.Debug($"First Init: {Inits.First().OriName}");

                string sBSTSeg = MappingPattern.Split('_')[11];
                if (sBSTSeg.StartsWith("SBST", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Scan";
                }
                string iniSBSTSeg = Inits.First().OriName.Split('_')[11];
                if (iniSBSTSeg.StartsWith("SBST", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Scan";
                }
            }

            return firstSeg;
        }

        internal bool IsJump()
        {
            return Regex.IsMatch(Search, "Jump", RegexOptions.IgnoreCase);
        }

        internal List<TestMethod> GetTestMethods()
        {
            var testMethods = new List<TestMethod>();
            string[] searchList = Search.Split(';');
            foreach (string search in searchList)
            {
                var testMethod = new TestMethod();
                string[] arr = search.Split(' ');
                if (arr.First().Equals("Jump", StringComparison.CurrentCultureIgnoreCase))
                {
                    testMethod.Name = arr.First();
                    testMethod.Arguments = arr.Length == 2 ? arr.Last() : "6";
                }
                else
                {
                    testMethod.Name = arr.First();
                }

                testMethods.Add(testMethod);
            }

            return testMethods;
        }

        internal Pin GetPinElementAt(int index)
        {
            if (string.IsNullOrEmpty(Order))
            {
                return Pins.FirstOrDefault(x => x.IsSearch) != null ? Pins.FirstOrDefault(x => x.IsSearch) : Pins.ElementAt(index);
            }

            var pinList = new List<Pin>();
            foreach (string pin in Order.Split(';'))
            {
                if (Pins.Exists(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                {
                    pinList.Add(Pins.Find(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)));
                }
            }

            if (pinList.Count > index)
            {
                return pinList.ElementAt(index);
            }

            return Pins.ElementAt(index);
        }

        internal string GetDfc()
        {
            const string pattern = @"\((?<value>[\w|,]+)\)";
            if (DataLoggingSetting.StartsWith("DFCList", StringComparison.CurrentCultureIgnoreCase))
            {
                string dfc = Regex.Match(DataLoggingSetting, pattern).Groups["value"].ToString();
                var arr = dfc.Split(',').Select(x => x.Trim()).Select(x => Regex.Replace(x, "mV$", "", RegexOptions.IgnoreCase)).ToList();
                var numbers = arr.Select(int.Parse).ToList();
                numbers.Sort();
                return "DFC" + string.Join("", numbers.Select(x => "M" + x));
            }
            if (DataLoggingSetting.StartsWith("DFCStep", StringComparison.CurrentCultureIgnoreCase))
            {
                string dfc = Regex.Match(DataLoggingSetting, pattern).Groups["value"].ToString();
                var arr = dfc.Split(',').Select(x => x.Trim()).Select(x => Regex.Replace(x, "mV$", "", RegexOptions.IgnoreCase)).ToList();
                if (arr.Count() == 2)
                {
                    return "DFCSize" + arr[0] + "Num" + arr[1];
                }
            }
            return "";
        }

        public string GetTestName()
        {
            List<string> items = new List<string>();
            items.Add(TestName);
            if (EnumDataLoggingSettingType == EnumDataLoggingSettingType.Dfc)
            {
                string dfc = GetDfc();
                if (!string.IsNullOrEmpty(dfc))
                {
                    items.Add(dfc);
                }
            }
            items.Add("CZ");
            items.Add(GetVoltageCategory());
            return string.Join("_", items);
        }

        public string GetTmpsFlowName()
        {
            var tmpsFlowName = new List<string> { "Flow_TMPS" };
            string low;
            string high;
            DataConvertor.GetTMPS_temperature(TempCondition, out low, out high);
            if (!string.IsNullOrEmpty(low))
            {
                tmpsFlowName.Add("L" + low);
            }
            if (!string.IsNullOrEmpty(high))
            {
                tmpsFlowName.Add("H" + high);
            }

            return string.Join("_", tmpsFlowName);
        }
        #endregion
    }
}
