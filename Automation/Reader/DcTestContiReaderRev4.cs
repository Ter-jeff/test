using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace Automation.Reader
{
    public class DcTestContiReaderRev4 : MySheetReader<DcTestContiSheet>
    {
        public int ContiInstanceIdx = -1;
        public int ContiJobIdx = -1;
        public int ContiDisableBinOutIdx = -1;
        public int ContiFailFlagIdx = -1;
        public int ContiSiteFlagIdx = -1;
        public int ContiEnableIdx = -1;
        public int ContiConditionIdx = -1;
        public int ContiModuleIdx = -1;
        public int ContiPinIdx = -1;
        public int ContiUnitsIdx = -1;

        public int ContiCp1LslIdx = -1;
        public int ContiCp1UslIdx = -1;
        public int ContiCp2LslIdx = -1;
        public int ContiCp2UslIdx = -1;
        public int ContiFt1LslIdx = -1;
        public int ContiFt1UslIdx = -1;
        public int ContiFt2LslIdx = -1;
        public int ContiFt2UslIdx = -1;
        public int ContiFt3LslIdx = -1;
        public int ContiFt3UslIdx = -1;

        public int ContiWlft1LslIdx = -1;
        public int ContiWlft1UslIdx = -1;
        public int ContiWlft2LslIdx = -1;
        public int ContiWlft2UslIdx = -1;

        private DcTestContiSheet ReadData()
        {
            var dcTestContiSheet = new DcTestContiSheet(ExcelWorksheet.Name);
            DcTestContiRow dcTestContiRow = null;
            for (int i = StartRow + 2; i <= EndRow; i++)
            {
                string instName = ExcelWorksheet.GetCellValue(i, ContiInstanceIdx).Trim();
                if (!string.IsNullOrEmpty(instName))
                {
                    dcTestContiRow = new DcTestContiRow
                    {
                        InstanceName = ExcelWorksheet.GetCellValue(i, ContiInstanceIdx).Trim().Replace(" ", ""),
                        Category = ExcelWorksheet.GetCellValue(i, ContiModuleIdx).Trim(),
                        PinGroup = ExcelWorksheet.GetCellValue(i, ContiPinIdx).Trim().Replace(" ", ""),
                        Condition = ExcelWorksheet.GetCellValue(i, ContiConditionIdx).Trim().Replace(" ", ""),
                        FailFlag = ExcelWorksheet.GetCellValue(i, ContiFailFlagIdx).Trim().Replace(" ", ""),
                        SiteFlag = ExcelWorksheet.GetCellValue(i, ContiSiteFlagIdx).Trim().Replace(" ", ""),
                        EnableWord = ExcelWorksheet.GetCellValue(i, ContiEnableIdx).Trim().Replace(" ", ""),
                        DisableBinOut = ExcelWorksheet.GetCellValue(i, ContiDisableBinOutIdx).Trim().Replace(" ", "")
                    };

                    ReadRow(ExcelWorksheet, i, dcTestContiRow);
                    dcTestContiSheet.DcTestContiRows.Add(dcTestContiRow);
                }
                else
                {
                    string pinGroup = ExcelWorksheet.GetCellValue(i, ContiPinIdx).Trim().Replace(" ", "");
                    string condition = ExcelWorksheet.GetCellValue(i, ContiConditionIdx).Trim().Replace(" ", "");
                    dcTestContiRow.PinGroup += !string.IsNullOrEmpty(pinGroup) ? ";" + pinGroup : "";
                    dcTestContiRow.Condition += !string.IsNullOrEmpty(condition) ? ";" + condition : "";
                    ReadRow(ExcelWorksheet, i, dcTestContiRow);
                }
            }

            return dcTestContiSheet;
        }

        private void ReadRow(ExcelWorksheet sheet, int rowIndex, DcTestContiRow contiRow)
        {
            string pinGrp = sheet.GetCellValue(rowIndex, ContiPinIdx).Trim().Replace(" ", "");
            string pinForceCondition = sheet.GetCellValue(rowIndex, ContiConditionIdx).Trim().Replace(" ", "");
            string jobStageStr = ExcelWorksheet.GetCellValue(rowIndex, ContiJobIdx).Trim().Replace(" ", "");
            string unitStr = ExcelWorksheet.GetCellValue(rowIndex, ContiUnitsIdx).Trim().Replace(" ", "");
            string failFlag = ExcelWorksheet.GetCellValue(rowIndex, ContiFailFlagIdx).Trim().Replace(" ", "");
            if (contiRow.JobNameList == null || contiRow.JobNameList.Count == 0)
            {
                contiRow.PinJobNameList.Add(pinGrp, jobStageStr.Split(',').ToList());
                contiRow.JobNameList = jobStageStr.Split(',').ToList();

            }
            else if (!string.IsNullOrEmpty(jobStageStr) && !string.IsNullOrEmpty(pinGrp))
            {
                contiRow.PinJobNameList.Add(pinGrp, jobStageStr.Split(',').ToList());
            }
            else
            {
                if (!contiRow.PinJobNameList.ContainsKey(pinGrp))
                {
                    contiRow.PinJobNameList.Add(pinGrp, contiRow.JobNameList);
                }
            }

            List<string> contiRowJob = string.IsNullOrEmpty(jobStageStr) ? contiRow.JobNameList : jobStageStr.Split(',').ToList();
            foreach (string jobHeader in contiRowJob)
            {
                var dcTestContiSheetLimit = new DcTestContiSheetLimit();
                switch (jobHeader.ToUpper())
                {
                    case DcTestContiRow.ConHeaderCp1:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderCp1;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiCp1LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiCp1UslIdx).Trim().Replace(" ", "");
                        break;

                    case DcTestContiRow.ConHeaderCp2:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderCp2;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiCp2LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiCp2UslIdx).Trim().Replace(" ", "");
                        break;

                    case DcTestContiRow.ConHeaderFt1:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderFt1;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiFt1LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiFt1UslIdx).Trim().Replace(" ", "");
                        break;

                    case DcTestContiRow.ConHeaderFt2:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderFt2;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiFt2LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiFt2UslIdx).Trim().Replace(" ", "");
                        break;
                    case DcTestContiRow.ConHeaderFt3:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderFt3;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiFt3LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiFt3UslIdx).Trim().Replace(" ", "");
                        break;
                    case DcTestContiRow.ConHeaderWlft1:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderWlft1;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiWlft1LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiWlft1UslIdx).Trim().Replace(" ", "");
                        break;
                    case DcTestContiRow.ConHeaderWlft2:
                        dcTestContiSheetLimit.LimitStage = DcTestContiRow.ConHeaderWlft2;
                        dcTestContiSheetLimit.LimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiWlft2LslIdx).Trim().Replace(" ", "");
                        dcTestContiSheetLimit.HiLimitValue = ExcelWorksheet.GetCellValue(rowIndex, ContiWlft2UslIdx).Trim().Replace(" ", "");
                        break;
                }
                dcTestContiSheetLimit.LimitHeader = pinGrp;
                dcTestContiSheetLimit.ForceConditionValue = pinForceCondition;
                dcTestContiSheetLimit.LimitUnit = unitStr;
                dcTestContiSheetLimit.FailFlag = failFlag;
                contiRow.Limits.Add(dcTestContiSheetLimit);
            }
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(2, i);

                GetGeneralHeaderIndex(header, i);

                string headerLimit = ExcelWorksheet.GetCellValue(1, i);

                GetLimitHeaderIndex(headerLimit, i);
            }
        }

        private void GetGeneralHeaderIndex(string header, int i)
        {
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderInstanceName))
            {
                ContiInstanceIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderJob))
            {
                ContiJobIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderDisableBinOut))
            {
                ContiDisableBinOutIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderPinGroup))
            {
                ContiPinIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderModuleUsage))
            {
                ContiModuleIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderCondition))
            {
                ContiConditionIdx = i;
            }
            if (header.Replace(" ", "").Equals(DcTestContiRow.ConHeaderFailFlag, StringComparison.OrdinalIgnoreCase))
            {
                ContiFailFlagIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderSiteFlag))
            {
                ContiSiteFlagIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderEnable))
            {
                ContiEnableIdx = i;
            }
            if (Regex.IsMatch(header, DcTestContiRow.ConHeaderSpecUnits))
            {
                ContiUnitsIdx = i;
            }
        }

        private void GetLimitHeaderIndex(string headerLimit, int i)
        {
            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderCp1))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiCp1LslIdx, ref ContiCp1UslIdx);
            }

            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderCp2))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiCp2LslIdx, ref ContiCp2UslIdx);
            }
            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderFt1))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiFt1LslIdx, ref ContiFt1UslIdx);
            }
            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderFt2))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiFt2LslIdx, ref ContiFt2UslIdx);
            }
            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderFt3))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiFt3LslIdx, ref ContiFt3UslIdx);
            }

            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderWlft1))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiWlft1LslIdx, ref ContiWlft1UslIdx);
            }

            if (Regex.IsMatch(headerLimit, DcTestContiRow.ConHeaderWlft2))
            {
                string limit = ExcelWorksheet.GetCellValue(2, i);

                GetLimitSideIndex(limit, i, ref ContiWlft2LslIdx, ref ContiWlft2UslIdx);
            }
        }

        private void GetLimitSideIndex(string limit, int i, ref int lslIdx, ref int uslIdx)
        {
            if (Regex.IsMatch(limit, "LSL"))
            {
                lslIdx = i;
            }

            if (Regex.IsMatch(limit, "USL"))
            {
                uslIdx = i;
            }
        }

        public override DcTestContiSheet ReadSheet(ExcelWorksheet worksheet)
        {
            ExcelWorksheet = worksheet;
            string sheetName = worksheet.Name;
            var sheet = new DcTestContiSheet(sheetName);

            if (!GetDimensions())
            {
                sheet.AddDimensionError();
                return sheet;
            }

            GetHeaderIndex();

            DcTestContiSheet contiSheet = ReadData();
            sheet.DcTestContiRows = contiSheet.DcTestContiRows;

            sheet.DicCategory = sheet.DcTestContiRows.GroupBy(p => p.SubBlock).ToDictionary(p => p.Key, p => p.ToList());


            return sheet;
        }
    }

    public class DcTestContiSheet : MySheet
    {
        public Dictionary<string, List<DcTestContiRow>> DicCategory;
        public List<DcTestContiRow> DcTestContiRows;

        public string PostFixed { get; }

        public DcTestContiSheet(string sheetName)
        {
            SheetName = sheetName;
            PostFixed = Regex.Match(sheetName, @"Continuity\s*_*(?<postName>\w+)", RegexOptions.IgnoreCase).Groups["postName"].Value;
            DcTestContiRows = new List<DcTestContiRow>();
            DicCategory = new Dictionary<string, List<DcTestContiRow>>();
        }
    }

    public class DcTestContiRow : MyRow
    {
        public const string ConHeaderInstanceName = "Instance";
        public const string ConHeaderJob = "Job";
        public const string ConHeaderDisableBinOut = "No Bin Out";
        public const string ConHeaderModuleUsage = "Module Usage";
        public const string ConHeaderPinGroup = "Pin Group";
        public const string ConHeaderCondition = "Condition";
        public const string ConHeaderFailFlag = "FailFlag";
        public const string ConHeaderSiteFlag = @"SiteFlag\s*\(per\s*site\)";
        public const string ConHeaderEnable = @"Enable\s*\(all\s*sites\)";
        public const string ConHeaderCp1 = "CP1";
        public const string ConHeaderCp2 = "CP2";
        public const string ConHeaderFt1 = "FT1";
        public const string ConHeaderFt2 = "FT2";
        public const string ConHeaderFt3 = "FT3";
        public const string ConHeaderWlft1 = "WLFT1";
        public const string ConHeaderWlft2 = "WLFT2";
        public const string ConHeaderSpecUnits = "Spec units";
        private const string Value = "Value";
        private const string Unit = "Unit";
        private const string OpenShortSource = @"Isource\w*\s*=\s*[+-]?(?<" + Value + @">(\d+[.])?\d+)(?<" + Unit + @">\w+)";
        private const string OpenShortSink = @"Isink\w*\s*=\s*[+-]?(?<" + Value + @">(\d+[.])?\d+)(?<" + Unit + @">\w+)";
        private const string PowerShortVforce = @"Vforce\w*\s*=\s*(?<" + Value + @">[+-]?(\d+[.])?\d+)(?<" + Unit + @">\w+)";
        private const string PowerShortIforce = @"Iforce\w*\s*=\s*(?<" + Value + @">[+-]?(\d+[.])?\d+)(?<" + Unit + @">\w+)";
        private const string DcCategoryCondition = @"DC\w*\s*=\s*(?<" + Value + @">\w+)";
        private const string Continuity = "Continuity(?!.*(Analog|PPMU))";
        private const string ContinuityAnalog = "Continuity.*Analog";
        private const string ContinuityPpmu = "Continuity.*PPMU";
        private const string PowerShort = @"Power\s*Short";
        private const string PowerSense = @"Power\s*Sense(?!.*Impedance)";
        private const string PowerImpedance = "Power.*Impedance";
        private const string GroundSense = @"Ground\s*Sense(?!.*Impedance)";
        private const string GroundImpedance = "Ground.*Impedance";
        private const string PowerOpenResistance = @"Power\s*Open.*Resistance";
        private const string WalkingZ = "Walking.*Z";
        private const string UserFunction = "UserFunction";
        private const string Opcode = "Opcode";
        private const string AutoZTtr = "AutoZTtr";
        private string _condition = "";

        private List<DcTestContiSheetLimit> _limits = new List<DcTestContiSheetLimit>();

        public string SubBlock = "";

        public string Category { set; get; } = "";

        public string InstanceName { set; get; } = "";

        public List<string> JobNameList { set; get; } = new List<string>();

        public string DisableBinOut { set; get; } = "";

        public Dictionary<string, List<string>> PinJobNameList { set; get; } = new Dictionary<string, List<string>>();

        public string PinGroup { set; get; } = "";

        public string SpecUnits { set; get; } = "";

        public string Condition
        {
            set
            {
                _condition = value;
                var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                string regexPattern = @"(?<key>[^=:;,\s]+)\s*[=:]\s*(?<value>(?:""[^""]*"")|[^;,\s]+)|(?<key>\w+)";
                MatchCollection matches = Regex.Matches(_condition, regexPattern);
                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        string inputKey = match.Groups["key"].Value.Trim();
                        string inputValue = match.Groups["value"].Success ? match.Groups["value"].Value.Trim().Trim('"').Trim() : "";
                        if (!match.Groups["value"].Success)
                        {
                            inputValue = inputKey;
                        }
                        if (result.ContainsKey(inputKey))
                        {
                            result[inputKey] += ";" + inputValue;
                        }
                        else
                        {
                            result.Add(inputKey, inputValue);
                        }
                    }
                }
                ConditionDict = result;
            }
            get
            {
                return _condition;
            }
        }

        public Dictionary<string, string> ConditionDict { get; private set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public string FailFlag { set; get; } = "";

        public string SiteFlag { set; get; } = "";

        public string EnableWord { set; get; } = "";

        public List<DcTestContiSheetLimit> Limits
        {
            set
            {
                _limits = value;
            }
            get
            {
                return _limits ?? (_limits = new List<DcTestContiSheetLimit>());
            }
        }

        public ContiType TestType
        {
            get
            {
                if (Regex.IsMatch(Category, Continuity, RegexOptions.IgnoreCase))
                {
                    return ContiType.OpenShort;
                }

                if (Regex.IsMatch(Category, PowerShort, RegexOptions.IgnoreCase))
                {
                    return ContiType.PowerShort;
                }

                if (Regex.IsMatch(Category, PowerSense, RegexOptions.IgnoreCase))
                {
                    return ContiType.PowerSense;
                }

                if (Regex.IsMatch(Category, PowerImpedance, RegexOptions.IgnoreCase))
                {
                    return ContiType.PowerImpedance;
                }

                if (Regex.IsMatch(Category, GroundSense, RegexOptions.IgnoreCase))
                {
                    return ContiType.GroundSense;
                }

                if (Regex.IsMatch(Category, GroundImpedance, RegexOptions.IgnoreCase))
                {
                    return ContiType.GroundImpedance;
                }

                if (Regex.IsMatch(Category, ContinuityAnalog, RegexOptions.IgnoreCase))
                {
                    return ContiType.ContiAnalog;
                }

                if (Regex.IsMatch(Category, ContinuityPpmu, RegexOptions.IgnoreCase))
                {
                    return ContiType.ContiPpmu;
                }

                if (Regex.IsMatch(Category, PowerOpenResistance, RegexOptions.IgnoreCase))
                {
                    return ContiType.Cres;
                }

                if (Regex.IsMatch(Category, WalkingZ, RegexOptions.IgnoreCase))
                {
                    return ContiType.WalkingZ;
                }

                if (Regex.IsMatch(Category, UserFunction, RegexOptions.IgnoreCase))
                {
                    return ContiType.UserFunction;
                }

                if (Category.Trim().Equals(Opcode, StringComparison.OrdinalIgnoreCase))
                {
                    return ContiType.Opcode;
                }

                if (Regex.IsMatch(Category, AutoZTtr, RegexOptions.IgnoreCase))
                {
                    return ContiType.AutoZTtr;
                }

                return ContiType.UnKnow;
            }
        }

        public Dictionary<string, string> GetForceConditions()
        {
            Dictionary<string, string> conditionList = new Dictionary<string, string>();

            foreach (string con in Condition.Replace(" ", "").Split(',').ToList())
            {
                if (string.IsNullOrEmpty(con))
                {
                    continue;
                }

                bool neg = false;
                string condition;
                string conditionValue;
                string unit;
                string type = con.Split('=')[0].Trim();

                //Get Condition
                if (Regex.IsMatch(con, OpenShortSource, RegexOptions.IgnoreCase))
                {
                    //ISource
                    conditionValue = Regex.Match(con, OpenShortSource, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, OpenShortSource, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(con, OpenShortSink, RegexOptions.IgnoreCase))
                {
                    //ISink
                    conditionValue = Regex.Match(con, OpenShortSink, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, OpenShortSink, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                    neg = true;
                }
                else if (Regex.IsMatch(con, PowerShortVforce, RegexOptions.IgnoreCase))
                {
                    //Vforce
                    conditionValue = Regex.Match(con, PowerShortVforce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, PowerShortVforce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(con, PowerShortIforce, RegexOptions.IgnoreCase))
                {
                    //Iforce
                    conditionValue = Regex.Match(con, PowerShortIforce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, PowerShortIforce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(con, DcCategoryCondition, RegexOptions.IgnoreCase))
                {
                    //DcCategory
                    conditionValue = Regex.Match(con, DcCategoryCondition, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = "";
                }
                else
                {
                    continue;
                }


                //Convert Unit
                if (unit != "" && string.Equals(type[0].ToString(), "I", StringComparison.OrdinalIgnoreCase))
                {
                    //Current
                    if (!conditionValue.TryCombineAmpere(unit, out condition))
                    {
                        condition = conditionValue;
                    }
                }
                else if (unit != "" && string.Equals(type[0].ToString(), "V", StringComparison.OrdinalIgnoreCase))
                {
                    //Volt
                    if (!conditionValue.TryCombineVolt(unit, out condition))
                    {
                        condition = conditionValue;
                    }
                }
                else if (unit == "" && conditionValue.Contains("Conti_", StringComparison.OrdinalIgnoreCase))
                {
                    //DC=Conti_X_X_X_X
                    condition = conditionValue;
                }
                else
                {
                    continue;
                    //condition = conditionValue;
                }

                string result;
                //Positive or negtive
                if (neg)
                {
                    result = "-" + condition;
                }
                else
                {
                    result = condition;
                }
                conditionList.Add(type.ToUpper(), result);
            }
            return conditionList;
        }

        public List<string> GetForceCondition()
        {
            string conditionValue;
            string unit;
            bool neg = false;
            string condition;
            ContiType type = TestType;
            var forceList = new List<string>();

            foreach (string subCondition in Condition.Split(';').Select(x => x.Trim()))
            {
                neg = false;
                //Get Condition
                if (Regex.IsMatch(subCondition, OpenShortSource, RegexOptions.IgnoreCase))
                {
                    //ISource
                    conditionValue = Regex.Match(subCondition, OpenShortSource, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(subCondition, OpenShortSource, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(subCondition, OpenShortSink, RegexOptions.IgnoreCase))
                {
                    //ISink
                    conditionValue = Regex.Match(subCondition, OpenShortSink, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(subCondition, OpenShortSink, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                    neg = true;
                }
                else if (Regex.IsMatch(subCondition, PowerShortVforce, RegexOptions.IgnoreCase))
                {
                    //Vforce
                    conditionValue = Regex.Match(subCondition, PowerShortVforce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(subCondition, PowerShortVforce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(subCondition, PowerShortIforce, RegexOptions.IgnoreCase))
                {
                    //Iforce
                    conditionValue = Regex.Match(subCondition, PowerShortIforce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(subCondition, PowerShortIforce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else
                {
                    forceList.Add("");
                    continue;
                }

                //Convert Unit
                if ((unit != "" && type == ContiType.OpenShort) || type == ContiType.ContiAnalog)
                {
                    //Current
                    if (!conditionValue.TryCombineAmpere(unit, out condition))
                    {
                        condition = conditionValue;
                    }
                }
                else if (unit != "" && type == ContiType.PowerShort)
                {
                    //Volt
                    if (conditionValue.TryCombineVolt(unit, out condition))
                    {
                        condition = conditionValue;
                    }
                }
                else
                {
                    condition = conditionValue;
                }

                //Positive or negtive
                if (neg)
                {
                    forceList.Add("-" + condition);
                }
                else
                {
                    forceList.Add(condition);
                }
            }
            return forceList;
        }

        public List<string> GetForceConditions(string forceConditions)
        {
            var results = new List<string>();

            foreach (string con in forceConditions.Replace(" ", "").Split(';', ',').ToList())
            {
                if (string.IsNullOrEmpty(con))
                {
                    continue;
                }

                bool neg = false;
                string condition;
                string conditionValue;
                string unit;
                string type = con.Split('=')[0].Trim();

                //Get Condition
                if (Regex.IsMatch(con, OpenShortSource, RegexOptions.IgnoreCase))
                {
                    //ISource
                    conditionValue = Regex.Match(con, OpenShortSource, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, OpenShortSource, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(con, OpenShortSink, RegexOptions.IgnoreCase))
                {
                    //ISink
                    conditionValue = Regex.Match(con, OpenShortSink, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, OpenShortSink, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                    neg = true;
                }
                else if (Regex.IsMatch(con, PowerShortVforce, RegexOptions.IgnoreCase))
                {
                    //Vforce
                    conditionValue = Regex.Match(con, PowerShortVforce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, PowerShortVforce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else if (Regex.IsMatch(con, PowerShortIforce, RegexOptions.IgnoreCase))
                {
                    //Iforce
                    conditionValue = Regex.Match(con, PowerShortIforce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                    unit = Regex.Match(con, PowerShortIforce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                }
                else
                {
                    continue;
                }


                //Convert Unit
                if (unit != "" && string.Equals(type[0].ToString(), "I", StringComparison.OrdinalIgnoreCase))
                {
                    //Current
                    if (!conditionValue.TryCombineAmpere(unit, out condition))
                    {
                        condition = conditionValue;
                    }
                }
                else if (unit != "" && string.Equals(type[0].ToString(), "V", StringComparison.OrdinalIgnoreCase))
                {
                    //Volt
                    if (!conditionValue.TryCombineVolt(unit, out condition))
                    {
                        condition = conditionValue;
                    }
                }
                else
                {
                    continue;
                    //condition = conditionValue;
                }

                string result;
                //Positive or negtive
                if (neg)
                {
                    result = "-" + condition;
                }
                else
                {
                    result = condition;
                }

                results.Add(result);

            }

            return results;
        }
    }

    public class DcTestContiSheetLimit
    {
        public string LimitStage { set; get; } = string.Empty;
        public string LimitHeader { set; get; } = string.Empty;
        public string LimitValue { set; get; } = string.Empty;
        public string LimitValueSecondary { set; get; } = string.Empty;
        public string LimitType { set; get; } = string.Empty;
        public string HiLimitValue { set; get; } = string.Empty;
        public string LoLimitValue { set; get; } = string.Empty;
        public string ForceConditionValue { set; get; } = string.Empty;
        public string LimitUnit { set; get; } = string.Empty;
        public string FailFlag { set; get; } = string.Empty;
    }

    public enum ContiType
    {
        PowerShort,
        OpenShort,
        PowerSense,
        PowerImpedance,
        GroundSense,
        GroundImpedance,
        ContiAnalog,
        Cres,
        WalkingZ,
        UserFunction,
        Opcode,
        UnKnow,
        ContiPpmu,
        AutoZTtr,
    }
}
