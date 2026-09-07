using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.IgxlProgramMappingLib;
using Cautogen.common.ReaderWriter.Reader;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.common.ReaderWriter.Reader.InputReader;
using Cautogen.Utility;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Utility;

using OfficeOpenXml;

using ExcelOperation = Cautogen.AutoCZ.CharPreProcessor.Utility.ExcelOperation;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader
{
    public class CharPlan : ExcelReader
    {
        public static HashSet<string> HardIpSheets = new HashSet<string>();
        public static RevisionSheet Revision = new RevisionSheet();
        public static Dictionary<string, List<Characterization>> CharPlanSheetDict;
        public static ManualAcSheet OptionalTimesetting;
        public static DigSrcReg DigSrcReg;
        public static AdaptiveCooling AdaptiveCooling;
        //public static List<SelSrmItem> SelMappingList = new List<SelSrmItem>();
        private readonly EmaMappingReader _emaMappingReader = new EmaMappingReader();
        private readonly HarvReader _harvReader = new HarvReader();
        private static readonly List<PinVarHeader> _pinVarDicList = new List<PinVarHeader>();
        private readonly bool _isRtosUseCmd = true;
        private bool _isMergeHlv = true;
        private readonly bool _isShmooHtoL = false;
        private bool _isMtd = false;
        public static bool IsContainsDashBoard = false;

        public static Dictionary<string, Characterization> AllcharList = new Dictionary<string, Characterization>();
        private static readonly Regex _eol = new Regex(@"[\n\r]", RegexOptions.Compiled);
        private readonly MappingResult _mappingResult;

        public CharPlan(MappingResult mappingResult, string filePath, bool rtosUseCmd, bool isMergeHlv = true, bool isHighToLow = false, Action callbackFunc = null)
            : base(filePath, callbackFunc)
        {
            _mappingResult = mappingResult;
            IsContainsDashBoard = false;
            HardIpSheets = new HashSet<string>();
            Revision = new RevisionSheet();
            CharPlanSheetDict = new Dictionary<string, List<Characterization>>();
            _isRtosUseCmd = rtosUseCmd;
            _isMergeHlv = isMergeHlv;
            _isShmooHtoL = isHighToLow;
            CallbackFuncList.Add(Revision.Check);
            //20230725 take out for new mapping rule
            //CallbackFuncList.Add(_UpdateDcAcCategory);
            //CallbackFuncList.Add(_FSelSRAMFlow);
        }

        public static void Reset()
        {
            HardIpSheets = new HashSet<string>();
            Revision = new RevisionSheet();
            CharPlanSheetDict = new Dictionary<string, List<Characterization>>();
            AllcharList = new Dictionary<string, Characterization>();
            OptionalTimesetting = null;
            DigSrcReg = null;
        }

        protected override void _Read(ExcelWorksheet sh)
        {
            if (sh.Dimension == null)
            {
                return;
            }

            MessageWriter.WriteMessage("Reading sheet: " + sh.Name + "...", EnumMessageLevel.Info);
            LogHelper.Info("Reading sheet: " + sh.Name + "...");
            if (!Regex.IsMatch(sh.Name, "EMADSSC_", RegexOptions.IgnoreCase))
            {
                switch (sh.Name.ToLower())
                {
                    //case "selsrm_mapping_table":
                    //    _selSrmReader.ReadMappingTable(sh);
                    //    SelMappingList = _selSrmReader.SrmItems;
                    //    break;

                    case "harv_mapping_table":
                        _harvReader.ReadMappingTable(sh);
                        break;

                    case "revision":
                        Revision.Read(sh);
                        break;

                    case "pinlist":
                        PinListReader.Read(sh);
                        break;

                    case "powermerge":
                        PowerMergeReader.Read(sh);
                        break;

                    case "timesettings":
                        OptionalTimesetting = new ManualAcSheet(sh);
                        break;

                    case "digsrcreg":
                        DigSrcReg = new DigSrcRegReader().Read(sh);
                        break;

                    case "adaptivecooling":
                        AdaptiveCooling = new AdaptiveCoolingReader().Read(sh);
                        break;

                    case "pattern_dashboard":
                        if (common.ReaderWriter.Reader.InputReader.PatternListInputReader.PatternList.Count == 0)
                        {
                            if (new common.ReaderWriter.Reader.InputReader.PatternListInputReader("").Read(sh))
                            {
                                IsContainsDashBoard = true;
                            }
                        }

                        break;

                    default:
                        ReadCharPlanSheet(sh);
                        break;
                }
            }
        }

        protected override void _ReadPreProcess(ExcelWorksheet sh)
        {
            try
            {
                if (Regex.IsMatch(sh.Name, "EMADSSC_", RegexOptions.IgnoreCase))
                {
                    UtilityMain.UtilityData.EmaMappingItems.AddRange(_emaMappingReader.ReadMappingTable(sh));
                }
            }
            catch (Exception)
            {
            }
        }

        private void _GetMapping(Characterization charRow)
        {
            if (_mappingResult == null)
            {
                return;
            }

            List<string> patternSplit = string.Join(",", charRow.AllPatterns.Values).Split(',').ToList();
            var patternList = new List<string>();
            PatternCell planPayload1Cell = charRow.PatternCellList.FirstOrDefault(x => x.Name == "PL1");
            string planPayload1 = "";
            if (planPayload1Cell != null)
            {
                planPayload1 = planPayload1Cell.PatternDefine.Split(':')[0].ToUpper();
            }

            foreach (string pat in patternSplit)
            {
                if (!string.IsNullOrEmpty(pat.Split(':')[0]))
                {
                    patternList.Add(pat.Split(':')[0].ToUpper());
                }
            }
            charRow.MappingKey = _mappingResult.CreateMappingCondition(patternList, planPayload1);
            charRow.MappingSpec = _mappingResult.GenMappingIntoInstance(charRow.MappingKey, charRow.IsHipRow());
        }

        private int _GetStartRow(ExcelWorksheet sheet)
        {
            var checkHeaderDict = new Dictionary<string, bool>()
            {
                { @"^USERDEF\d+$", false },
                { @"^INIT\d+$", false },
                { @"^PAYLOAD\d+$", false }
            };
            /* Return the row number of Column A == "Use/Not Use" */
            for (int row = 1; row <= sheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    if (sheet.Cells[row, col].Value == null)
                    {
                        continue;
                    }

                    string cell = sheet.Cells[row, col].Value.ToString().Trim();
                    if (string.IsNullOrEmpty(cell))
                    {
                        continue;
                    }

                    var checkHeaders = checkHeaderDict.Keys.ToList();
                    foreach (string checkHeader in checkHeaders)
                    {
                        if (Regex.IsMatch(cell, checkHeader, RegexOptions.IgnoreCase))
                        {
                            checkHeaderDict[checkHeader] = true;
                        }
                    }
                }

                if (checkHeaderDict.Values.All(x => x))
                {
                    return row;
                }
                else
                {
                    var checkHeaders = checkHeaderDict.Keys.ToList();
                    foreach (string checkHeader in checkHeaders)
                    {
                        checkHeaderDict[checkHeader] = false;
                    }
                }
            }
            return sheet.Dimension.End.Row;
        }

        public bool ReadCharPlanSheet(ExcelWorksheet sheet)
        {
            // reset
            var headerOrder = new Dictionary<string, int>();
            var shmooList = new List<ShmooSpec>();
            var charList = new List<Characterization>();

            // early return
            int startRow = _GetStartRow(sheet);
            int endRow = sheet.Dimension.End.Row;
            if (startRow == endRow || startRow == 1)
            {
                return false;
            }

            // update header
            _UpdateHeaderOrder(sheet, startRow, shmooList, headerOrder);
            if (!_IsCheckHeaderFormatPass(sheet.Name, startRow, headerOrder))
            {
                return false;
            }

            if (headerOrder.ContainsKey("die"))
            {
                _isMergeHlv = false;
                _isMtd = true;

            }
            // read content of each row
            try
            {
                for (int j = startRow + 1; j <= endRow; j++)
                {
                    ProcessCharPlanRow(sheet, j, headerOrder, shmooList, charList);
                }
            }
            catch (Exception e)
            {
                MessageWriter.WriteMessage("Failed when reading " + sheet.Name + ", please check! " + e.Data + e.Message, EnumMessageLevel.Error);
            }

            FinalizeCharList(sheet, charList);
            return true;
        }

        private void ProcessCharPlanRow(ExcelWorksheet sheet, int j, Dictionary<string, int> headerOrder, List<ShmooSpec> shmooList, List<Characterization> charList)
        {
            Characterization charRow = BuildCharRow(sheet, j, headerOrder);

            if (charRow.Use == "")
            {
                if (!string.IsNullOrEmpty(charRow.TpName))
                {
                    const string outString = "Use/Not Use of this Char Item is empty, please check";
                    ErrorManager.AddWarning(ErrorType.EmptyUse, charRow.SheetName, charRow.RowNum,
                        charRow.ColNum("use/notuse"), "use", outString);
                    foreach (int col in charRow.ColNum("use/notuse"))
                    {
                        ErrorReportManager.AddError(CharErrorType.W_EmptyUse_01, charRow.SheetName, charRow.RowNum, col, []);
                    }
                }
                return;
            }
            _GetPatternHeaders(sheet, j, charRow, headerOrder);
            _UpdateExtraPatterns(sheet, j, charRow, headerOrder);
            _GetMapping(charRow);
            // check if it's HIP row
            if (charRow.IsHipRow())
            {
                HardIpSheets.Add(sheet.Name.ToUpper());
            }
            else if (charRow.UserDef2.ToUpper().Equals("AMP"))
            {
                string outString = charRow.UserDef1 + " is not use for hardip.";
                ErrorManager.AddWarning(ErrorType.WrongUserdef1, charRow.SheetName, charRow.RowNum, charRow.ColNum("USERDEF1"), charRow.Use, outString);
            }

            //20230725 take out for new mapping rule
            //if (UtilityMain.UtilityData.InputParam.DefFile != "" && charRow.Category != "" && !InputDefReader.BlockMapDict.ContainsKey(charRow.Category) && !UtilityMain.UtilityData.MissingCategory.Contains(charRow.Category))
            //{
            //    UtilityMain.UtilityData.MissingCategory.Add(charRow.Category);
            //    var outString = "Missing Category in Char_Input_Def: " + charRow.Category;
            //    ErrorManager.AddError(ErrorType.MissingCategory, charRow.SheetName, charRow.RowNum, charRow.ColNum("category"), charRow.Use, outString);
            //}

            bool isVoltagetable = AssignOtherSupplies(sheet, j, charRow, headerOrder);
            if (!isVoltagetable)
            {
                _UpdateDcSelectorFromOtherSupplies(sheet, j, charRow, "othersupplies", _GetCell(sheet, j, "othersupplies", headerOrder).Trim().ToUpper());
            }
            else
            {
                _UpdateDcSelectorFromOtherSupplies(sheet, j, charRow, "voltagetable", _GetCell(sheet, j, "voltagetable", headerOrder).Trim().ToUpper());
            }
            _UpdateShmoo(sheet, charRow, j, shmooList, _isShmooHtoL);
            _UpdatePinSweep(sheet, j, charRow);
            _CheckTestNameFormula(sheet, charRow);
            if (!_isMtd && !_CheckNoDuplicateTestNamePass(charRow))
            {
                return;
            }

            if (_isMergeHlv && _HasMergeLhv(charRow, charList))
            {
                return;
            }

            foreach (Characterization item in _ExtendCharItemWithAC(charRow))
            {
                charList.Add(item);
            }

            if (!UtilityMain.UtilityData.InputParam.IsSplitSgmt32 || !_IsHacSgmt32(charRow))
            {
                return;
            }

            var spliteCharRow = new Characterization(charRow);
            spliteCharRow.SpliteSgmt32();
            charList.Add(spliteCharRow);
        }

        private Characterization BuildCharRow(ExcelWorksheet sheet, int j, Dictionary<string, int> headerOrder)
        {
            return new Characterization
            {
                SheetName = sheet.Name,
                RowNum = j,
                ColumnDic = headerOrder,
                Use = _GetCell(sheet, j, "use/notuse", headerOrder),
                Enableword = _GetCell(sheet, j, "enableword", headerOrder),
                FailFlag = _GetCell(sheet, j, "failflag", headerOrder),
                Update = _GetCell(sheet, j, "updated", headerOrder),
                Priority = _GetCell(sheet, j, "priority", headerOrder),
                UserDef1 = _GetCell(sheet, j, "userdef1", headerOrder),
                UserDef2 = _GetCell(sheet, j, "userdef2", headerOrder),
                UserDef3 = _GetCell(sheet, j, "userdef3", headerOrder),
                UserDef4 = _GetCell(sheet, j, "userdef4", headerOrder),
                UserDef5 = _GetCell(sheet, j, "userdef5", headerOrder),
                UserDef6 = _GetCell(sheet, j, "userdef6", headerOrder),
                UserDef7 = _GetCell(sheet, j, "userdef7", headerOrder),
                UserDef8 = _GetCell(sheet, j, "userdef8", headerOrder),
                UserDef9 = _GetCell(sheet, j, "userdef9", headerOrder),
                OriTpName = _GetCell(sheet, j, "tpname(teusesuffix)", headerOrder),
                TpName = _GetCell(sheet, j, "tpname(teusesuffix)", headerOrder),  // will be modify later
                Group = _GetCell(sheet, j, "group", headerOrder),
                Init1 = _GetCell(sheet, j, "init1", headerOrder),
                Init2 = _GetCell(sheet, j, "init2", headerOrder),
                Init3 = _GetCell(sheet, j, "init3", headerOrder),
                Init4 = _GetCell(sheet, j, "init4", headerOrder),
                Init5 = _GetCell(sheet, j, "init5", headerOrder),
                Init6 = _GetCell(sheet, j, "init6", headerOrder),
                Init7 = _GetCell(sheet, j, "init7", headerOrder),
                Init8 = _GetCell(sheet, j, "init8", headerOrder),
                Init9 = _GetCell(sheet, j, "init9", headerOrder),
                Init10 = _GetCell(sheet, j, "init10", headerOrder),
                Payload1 = _GetCell(sheet, j, "payload1", headerOrder),
                Payload2 = _GetCell(sheet, j, "payload2", headerOrder),
                Payload3 = _GetCell(sheet, j, "payload3", headerOrder),
                Payload4 = _GetCell(sheet, j, "payload4", headerOrder),
                Payload5 = _GetCell(sheet, j, "payload5", headerOrder),
                //ColumnNum = headerOrder["payload1"],
                IpUse1 = _GetCell(sheet, j, "testitems", headerOrder),
                IpUse2 = _GetCell(sheet, j, "description", headerOrder),
                IpUse3 = UtilityFunction.NormalizeUnit(_GetCell(sheet, j, "usl", headerOrder)),
                IpUse4 = UtilityFunction.NormalizeUnit(_GetCell(sheet, j, "lsl", headerOrder)),
                IpUse5 = _GetCell(sheet, j, "units", headerOrder),
                Category = _GetCell(sheet, j, "category", headerOrder).ToUpper(),
                YAxis = _GetCell(sheet, j, "y-axis", headerOrder),
                Retention = _RetentionGuardBandPlus(UtilityFunction.NormalizeUnit(_GetCell(sheet, j, new List<string>() { "retentiontime", "retentiontime:guardband" }, headerOrder))),
                TestCondition = _GetCell(sheet, j, "testcondition", headerOrder),
                PowerRunScenario = _GetCell(sheet, j, "powerrunscenario", headerOrder),
                Ttr = _GetCell(sheet, j, "ttr", headerOrder),
                Search = _GetCell(sheet, j, "method", headerOrder),
                Nominal = UtilityFunction.NormalizeUnit(_GetCell(sheet, j, "voltage", headerOrder)),
                Htol = _GetCell(sheet, j, "htol", headerOrder),
                IsUseRtosCmd = _isRtosUseCmd && _GetCell(sheet, j, "category", headerOrder).ToLower().Contains("rtos"),
                SuspendDatalog = _GetCell(sheet, j, "suspenddatalog", headerOrder),
                TimeSet = _GetCell(sheet, j, "timeset", headerOrder),
                ShiftFreq = _GetCell(sheet, j, "shiftfrequency", headerOrder),
                HarvFstp = _GetCell(sheet, j, "harv_fstp", headerOrder),
                SiteFlag = _GetCell(sheet, j, "siteflag", headerOrder),
                FailInfo = _GetCell(sheet, j, "failinfo", headerOrder),
                Dfc = _GetCell(sheet, j, "dfc", headerOrder),
                Environment = _GetCell(sheet, j, "environment", headerOrder),
                Burst = _GetCell(sheet, j, "updated", headerOrder),
                DigSrcShiftOrder = _GetCell(sheet, j, "digsrcshiftorder", headerOrder).Equals("LSB2MSB", StringComparison.OrdinalIgnoreCase),
                BypassShmooHole = _GetCell(sheet, j, "bypass_shmoohole", headerOrder),
                RetentionRamp = _GetCell(sheet, j, "retentionrampstep:time", headerOrder),
                OneTimeInit = _GetCell(sheet, j, "One_Time_INIT", headerOrder),
                ApplyVoltageFromBinCut = _GetCell(sheet, j, "ApplyVoltageFromBinCut", headerOrder),
                UserFunction = _GetCell(sheet, j, "UserFunction", headerOrder),
                HarvestPinGrpOtherFail = _GetCell(sheet, j, "HarvestPinGrpOtherFail", headerOrder),
                EnableCoreHarvest = _GetCell(sheet, j, "EnableCoreHarvest", headerOrder),
                EnableCoreMask = _GetCell(sheet, j, "EnableCoreMask", headerOrder),
                PinGrpSpecifyMask = _GetCell(sheet, j, "PinGrp_SpecifyMask", headerOrder),
                SsnSpecifyMask = _GetCell(sheet, j, "SSN_SpecifyMask", headerOrder),
                AdaptiveCooling = _GetCell(sheet, j, "AdaptiveCooling", headerOrder),
                //Adding for WAT Die infomation
                StageCp1 = _GetCell(sheet, j, "cp1", headerOrder),
                StageCp2 = _GetCell(sheet, j, "cp2", headerOrder),
                StageFt1 = _GetCell(sheet, j, "ft1", headerOrder),
                StageFt2 = _GetCell(sheet, j, "ft2", headerOrder),
                Die = _GetCell(sheet, j, "DIE", headerOrder),

            };
        }

        private bool AssignOtherSupplies(ExcelWorksheet sheet, int j, Characterization charRow, Dictionary<string, int> headerOrder)
        {
            bool isVoltagetable = false;
            if (HardIpSheets.Contains(sheet.Name.ToUpper()) && !charRow.IsFuncRow())
            {
                charRow.OtherSupplies = UtilityMain.UtilityFunction.ConvertVoltage(charRow.UserDef3.ToUpper());
            }
            else if (!string.IsNullOrEmpty(_GetCell(sheet, j, "othersupplies", headerOrder)))
            {
                charRow.OtherSupplies = _GetCell(sheet, j, "othersupplies", headerOrder);
            }
            else //C651 Toby want to use "Voltage Table" as header unstead of "Other Supplies, 2021/9/14"
            {
                isVoltagetable = true;
                charRow.OtherSupplies = _GetCell(sheet, j, "voltagetable", headerOrder);
            }
            return isVoltagetable;
        }

        private void FinalizeCharList(ExcelWorksheet sheet, List<Characterization> charList)
        {
            if (charList.Count > 0)
            {
                CharPlanSheetDict[sheet.Name] = charList;
                _GetDigSrc(charList);
                if (!_isMtd)
                {
                    _GetDigSrcForNewTChar(charList);
                }
                else
                {
                    _GetDigSrcForMTDChar(charList);
                }
            }
        }

        private static string _RetentionGuardBandPlus(string str)
        {
            var resultList = new List<string>();
            string[] patternsRetention = str.Split(',');
            foreach (string patternRetention in patternsRetention)
            {
                string[] splitSegments = patternRetention.Split(':');
                if (splitSegments.Length > 2)
                {
                    if (double.TryParse(splitSegments[2], out double parseValue))
                    {
                        if (parseValue >= 0)
                        {
                            splitSegments[2] = "+" + parseValue;
                        }
                    }
                }
                resultList.Add(string.Join(":", splitSegments));
            }
            return string.Join(",", resultList);
        }

        private static void _GetPatternHeaders(ExcelWorksheet ws, int row, Characterization charItem, Dictionary<string, int> headerDictionary)
        {
            var patternCellList = new List<PatternCell>();
            //For init patterns
            List<string> keys = headerDictionary.Keys.ToList().FindAll(p => Regex.IsMatch(p, @"init(?<Num>\d+)", RegexOptions.IgnoreCase));
            var initDic = new SortedDictionary<int, string>();
            foreach (string initKey in keys)
            {
                int.TryParse(Regex.Match(initKey, @"init(?<Num>\d+)", RegexOptions.IgnoreCase).Groups["Num"].ToString(), out int num);
                initDic.Add(num, initKey);
            }
            foreach (KeyValuePair<int, string> init in initDic.Where(init => _GetCell(ws, row, init.Value, headerDictionary) != ""))
            {
                string patternDefine = _GetCell(ws, row, init.Value, headerDictionary);
                if (!string.IsNullOrEmpty(patternDefine))
                {
                    patternCellList.Add(new PatternCell
                    {
                        Header = init.Value,
                        Name = "INIT" + init.Key,
                        ColIndex = headerDictionary[init.Value],
                        PatternDefine = _GetCell(ws, row, init.Value, headerDictionary),
                    });
                }
            }
            //For pl patterns
            keys = headerDictionary.Keys.ToList().FindAll(p => Regex.IsMatch(p, @"payload(?<Num>\d+)", RegexOptions.IgnoreCase));
            var plDic = new SortedDictionary<int, string>();
            foreach (string plKey in keys)
            {
                int.TryParse(Regex.Match(plKey, @"payload(?<Num>\d+)", RegexOptions.IgnoreCase).Groups["Num"].ToString(), out int num);
                plDic.Add(num, plKey);
            }
            foreach (KeyValuePair<int, string> pl in plDic.Where(pl => _GetCell(ws, row, pl.Value, headerDictionary) != ""))
            {
                string patternDefine = _GetCell(ws, row, pl.Value, headerDictionary);
                if (!string.IsNullOrEmpty(patternDefine))
                {
                    patternCellList.Add(new PatternCell
                    {
                        Header = pl.Value,
                        Name = "PL" + pl.Key,
                        ColIndex = headerDictionary[pl.Value],
                        PatternDefine = _GetCell(ws, row, pl.Value, headerDictionary),
                    });
                }
            }
            //Save 
            charItem.PatternCellList = patternCellList;
        }

        private static bool _IsHacSgmt32(Characterization charRow)
        {
            /* check if sgmt bit legnt == 32*/
            if (charRow.UserDef1.ToUpper() != "HAC")
            {
                return false;
            }

            string pat = charRow.Payload1.ToUpper();
            if (!UtilityMain.UtilityData.PatInfoDict.TryGetValue(pat, out HardIpReference patInfo))
            {
                return false;
            }

            Dictionary<string, int> capBitLenDict = _GetCapBitLenDict(patInfo);
            string sgmtName = charRow.UserDef6.ToUpper();
            if (!capBitLenDict.TryGetValue(sgmtName, out int value))
            {
                return false;
            }

            return value == 32;
        }

        private static Dictionary<string, int> _GetCapBitLenDict(HardIpReference patInfo)
        {
            var capBitLenDict = new Dictionary<string, int>();
            foreach (string[] sgmtBitlenTuple in patInfo.CapBitStr.Split('+')
                .Select(capBitStr => capBitStr.Split('_')).Where(sgmtBitlenTuple => sgmtBitlenTuple.Length == 2))
            {
                int.TryParse(sgmtBitlenTuple[1], out int bitLen);
                capBitLenDict[sgmtBitlenTuple[0].ToUpper()] = bitLen;
            }
            return capBitLenDict;
        }

        private static void _UpdateHeaderOrder(ExcelWorksheet sheet, int startRow, ICollection<ShmooSpec> shmooList, IDictionary<string, int> headerOrder)
        {
            int startCol = sheet.Dimension.Start.Column;
            int endCol = sheet.Dimension.End.Column;

            _pinVarDicList.Clear();

            // loop each col for the start row (usually row 2)
            for (int col = startCol; col <= endCol; col++)
            {
                // get header cell value
                string header;
                if (sheet.Cells[startRow, col].Value != null)
                {
                    header = sheet.Cells[startRow, col].Value.ToString().Trim();
                }
                else
                {
                    continue;
                }

                // empty header checker
                if (header.Trim() == "")
                {
                    string address = ExcelCellBase.GetAddress(startRow, col);
                    ErrorManager.AddError(ErrorType.EmptyHeader, sheet.Name, startRow, new List<int> { col },
                        "use", $"EmptyColumn Need Check:{address}");
                    foreach (int column in new List<int> { col })
                    {
                        ErrorReportManager.AddError(CharErrorType.E_EmptyHeader_01, sheet.Name, startRow, column, [address]);
                    }
                    continue;
                }

                // headerOrder = {header: col_index}
                headerOrder[header.Replace(" ", "").Replace("_", "").ToLower().Split('[').First()] = col;

                // collect (pin | variable) col information
                if (header.Equals("Pin", StringComparison.OrdinalIgnoreCase) ||
                    header.Equals("variable", StringComparison.OrdinalIgnoreCase))
                {
                    var pinVarIndex = new PinVarHeader
                    {
                        IsVar = header.Equals("variable", StringComparison.OrdinalIgnoreCase),
                        PinVar = col++
                    };

                    _pinVarDicList.Add(pinVarIndex);
                    string varheader;
                    while (!string.IsNullOrEmpty(varheader = (string)sheet.Cells[startRow, col].Value))
                    {
                        if (pinVarIndex.Start == -1 && varheader.Equals("start", StringComparison.OrdinalIgnoreCase))
                        {
                            pinVarIndex.Start = col++;
                        }
                        else if (pinVarIndex.Stop == -1 && varheader.Equals("stop", StringComparison.OrdinalIgnoreCase))
                        {
                            pinVarIndex.Stop = col++;
                        }
                        else if (pinVarIndex.Step == -1 && varheader.Equals("step", StringComparison.OrdinalIgnoreCase))
                        {
                            pinVarIndex.Step = col++;
                        }
                        else if (pinVarIndex.Type == -1 && varheader.Equals("type", StringComparison.OrdinalIgnoreCase))
                        {
                            pinVarIndex.Type = col++;
                        }
                        else if (pinVarIndex.Order == -1 && varheader.Equals("order", StringComparison.OrdinalIgnoreCase))
                        {
                            pinVarIndex.Order = col++;
                        }
                        else
                        {
                            col--;
                            break;
                        }
                    }
                    continue;
                }

                // add start power domain or pin to the shmoo List {Name: powerDomainName, ColIndex: col_index}
                if (header == "Start")
                {
                    object powerDomainName = sheet.Cells[startRow - 1, col].Value;

                    if (powerDomainName == null)
                    {
                        continue;
                    }

                    string powerDomainNameString = powerDomainName.ToString();
                    var shmooItem = new ShmooSpec
                    {
                        IsValt = Regex.IsMatch(powerDomainNameString, @"\s*valt", RegexOptions.IgnoreCase),
                        IsVret = Regex.IsMatch(powerDomainNameString, @"\s*vret", RegexOptions.IgnoreCase),
                        Name = Regex.Replace(powerDomainNameString, @"\s*valt", "", RegexOptions.IgnoreCase).Trim().Replace("_", ""),
                        ColIndex = col
                    };
                    if (sheet.Cells[startRow, col + 3].Text.Equals("order", StringComparison.OrdinalIgnoreCase) ||
                        sheet.Cells[startRow, col + 4].Text.Equals("order", StringComparison.OrdinalIgnoreCase))//power/pin/variable condition
                    {
                        shmooItem.IsContainsOrder = true;
                    }

                    shmooList.Add(shmooItem);

                    _PowerDomainNameChecker(powerDomainNameString, sheet.Name, startRow, col);
                }
            }
        }

        private static void _PowerDomainNameChecker(string powerDomainNameString, string sheetName, int row, int col)
        {
            // sweep condition should apply to pin group for vdd_low and vdd_fixed, instead of indivisual of them
            bool isLowOrFix = Regex.IsMatch(powerDomainNameString, "VDD_*((LOW)|(FIXED))", RegexOptions.IgnoreCase);
            bool isGroup = Regex.IsMatch(powerDomainNameString, "group", RegexOptions.IgnoreCase);

            if (isLowOrFix && !isGroup)
            {
                ErrorManager.AddWarning(
                    ErrorType.PowerGroupApply, sheetName, row - 1,
                    new List<int> { col }, "use", $"Suggest tie all {powerDomainNameString} related power together.");
                foreach (int column in new List<int> { col })
                {
                    ErrorReportManager.AddError(CharErrorType.W_PowerGroupApply_01, sheetName, row - 1, column, [powerDomainNameString]);
                }
            }
        }

        private static string _GetFrcNameByPlan(string name)
        {
            name = name.Replace(" ", "").ToUpper();

            foreach (string frc in UtilityMain.UtilityData.FrcList)
            {
                string frcUpper = frc.Trim().Replace("_", "").ToUpper();

                if (name == frcUpper)
                {
                    return frc.Trim().Replace("_", "");
                }

                if (frcUpper.StartsWith(Regex.Replace(name, "VAR$", ""), StringComparison.OrdinalIgnoreCase))
                {
                    return frc.Trim().Replace("_", "");
                }

                if (frcUpper.StartsWith(Regex.Replace(name, "FREQVAR$", ""), StringComparison.OrdinalIgnoreCase))
                {
                    return frc.Trim().Replace("_", "");
                }

                if (frcUpper.StartsWith(Regex.Replace(name, "DIFFFREQVAR$", ""), StringComparison.OrdinalIgnoreCase))
                {
                    return frc.Trim().Replace("_", "");
                }
            }
            return "";
        }

        private static void _UpdateShmoo(ExcelWorksheet sheet, Characterization charRow, int j, IEnumerable<ShmooSpec> shmooList, bool isShmooHtoL)
        {
            foreach (ShmooSpec shmoo in shmooList)
            {
                var shmooSpec = new ShmooSpec
                {
                    Name = shmoo.Name,
                    Type = Regex.IsMatch(shmoo.Name, "^vdd", RegexOptions.IgnoreCase) ? "VDD" : "AC Spec",
                    ColIndex = shmoo.ColIndex,
                    Start = UtilityFunction.NormalizeUnit(ExcelOperation.GetCellValue(sheet, j, shmoo.ColIndex)),
                    Stop = UtilityFunction.NormalizeUnit(ExcelOperation.GetCellValue(sheet, j, shmoo.ColIndex + 1)),
                    Step = UtilityFunction.NormalizeUnit(ExcelOperation.GetCellValue(sheet, j, shmoo.ColIndex + 2)),
                    Order = shmoo.IsContainsOrder ? ExcelOperation.GetCellValue(sheet, j, shmoo.ColIndex + 3) : "",
                    IsValt = shmoo.IsValt,
                    IsVret = shmoo.IsVret
                };

                if (shmooSpec.Type == "AC Spec")
                {
                    string prefix = _GetFrcNameByPlan(shmoo.Name);
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        shmooSpec.Type = "FRC";
                        shmooSpec.Name = prefix.Trim().Replace("_", "");
                    }
                }

                if (shmooSpec.Stop == "")
                {
                    shmooSpec.Stop = shmooSpec.Start;
                }

                if (isShmooHtoL && shmooSpec.Type.Equals("VDD"))
                {
                    if (double.Parse(shmooSpec.Stop) > double.Parse(shmooSpec.Start))
                    {
                        string tempValue = shmooSpec.Stop;
                        shmooSpec.Stop = shmooSpec.Start;
                        shmooSpec.Start = tempValue;
                    }
                }

                if (shmooSpec.Order.IndexOf("Y", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    charRow.PowerSupplyY.Add(shmooSpec);
                }
                else if (shmooSpec.Order.IndexOf("Z", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    charRow.PowerSupplyZ.Add(shmooSpec);
                }
                else
                {
                    charRow.PowerSupplyX.Add(shmooSpec);
                }
            }

            // move the UD3 Shmoo to the first of PowerSupplyX
            int ud3Index = charRow.PowerSupplyX
                .FindIndex(shmoo => shmoo.Name.Equals("VDD" + Regex.Replace(charRow.UserDef3, "^VDD", "")));
            if (ud3Index == -1)
            {
                return;
            }

            ShmooSpec ud3Shmoo = charRow.PowerSupplyX[ud3Index];
            charRow.PowerSupplyX.RemoveAt(ud3Index);
            charRow.PowerSupplyX.Insert(0, ud3Shmoo);
        }

        private static void _UpdatePinSweep(ExcelWorksheet sheet, int j, Characterization charRow)
        {
            bool isSweepOnFreq = false;
            foreach (PinVarHeader pinVarItem in _pinVarDicList)
            {
                string pinName = ExcelOperation.GetCellValue(sheet, j, pinVarItem.PinVar);
                if (string.IsNullOrEmpty(pinName))
                {
                    continue;
                }

                var pin = new Pin
                {
                    PinName = pinName,
                    PlanIndex = pinVarItem.PinVar,
                    PinType = pinVarItem.Type != -1 ? ExcelOperation.GetCellValue(sheet, j, pinVarItem.Type) : "AC Spec",
                    Start = pinVarItem.Start != -1 ? UtilityFunction.NormalizeUnit(ExcelOperation.GetCellValue(sheet, j, pinVarItem.Start)) : "",
                    Stop = pinVarItem.Stop != -1 ? UtilityFunction.NormalizeUnit(ExcelOperation.GetCellValue(sheet, j, pinVarItem.Stop)) : "",
                    Step = pinVarItem.Step != -1 ? UtilityFunction.NormalizeUnit(ExcelOperation.GetCellValue(sheet, j, pinVarItem.Step)) : "",
                    Order = pinVarItem.Order != -1 ? ExcelOperation.GetCellValue(sheet, j, pinVarItem.Order) : ""
                };
                if (pin.Start.Equals(pin.Stop) && pin.PinType.Equals("AC Spec", StringComparison.OrdinalIgnoreCase))
                {
                    pin.Step = "1000000"; //-> for one point frequency running
                }

                if (pin.PinType.ToUpper() == "AC SPEC")
                {
                    isSweepOnFreq = true;  // todo: need other method to decide if it's sweep on freq
                }

                if (pin.PinType.ToUpper() == "VCM")
                {
                    pin.PinType = "VICM";  // pre Chihome's request
                }

                if (pin.Start == pin.Stop && pin.PinType.ToUpper() != "AC SPEC")
                {
                    pin.PinCondition = pin.Start;
                }

                charRow.Pins.Add(pin);
            }

            _UpdateSearchMethod(charRow, isSweepOnFreq);
        }

        private static void _UpdateSearchMethod(Characterization charRow, bool isSweepOnFreq)
        {
            // pre check search method should be liner for sweep on freq char row
            if (isSweepOnFreq && charRow.Search.ToLower() != "linear")
            {
                ErrorManager.AddError(ErrorType.ErrorMethod, charRow.SheetName, charRow.RowNum, charRow.ColNum("method"), charRow.Use, "Error Search Method");
                foreach (int col in charRow.ColNum("method"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_ErrorMethod_01, charRow.SheetName, charRow.RowNum, col, []);
                }
            }

            // using default search method if not specified
            if (charRow.Search == "" || charRow.Search == "X")
            {
                charRow.Search = isSweepOnFreq ? "Linear" : "Jump";
            }
        }

        private static void _UpdateDcSelectorFromOtherSupplies(ExcelWorksheet sheet, int j, Characterization charRow, string header, string otherSupply)
        {
            switch (otherSupply.Split(':').First().Split(' ').Last().Trim())
            {
                case "NV":
                    charRow.DcSelector = "Typ";
                    break;
                case "HV":
                    charRow.DcSelector = "Max";
                    break;
                case "LV":
                    charRow.DcSelector = "Min";
                    break;
                default:
                    string outString = "missing or wrong other supplies in " + sheet.Name + " Row: " + j;
                    ErrorManager.AddError(ErrorType.NoOtherSupplies, charRow.SheetName, j, charRow.ColNum(header), charRow.Use, outString);
                    foreach (int col in charRow.ColNum(header))
                    {
                        ErrorReportManager.AddError(CharErrorType.E_NoOtherSupplies_01, charRow.SheetName, j, col, [sheet.Name, $"{j}"]);
                    }
                    break;
            }
        }

        private static string _GetCell(ExcelWorksheet sheet, int j, string fieldName, Dictionary<string, int> headerOrder)
        {
            fieldName = fieldName.Replace("_", "").ToLower();
            return headerOrder.Keys.Contains(fieldName)
                ? _eol.Replace(ExcelOperation.GetCellValue(sheet, j, headerOrder[fieldName]), "")
                : "";
        }

        private static string _GetCell(ExcelWorksheet sheet, int j, List<string> fieldNameList, Dictionary<string, int> headerOrder)
        {
            foreach (string fieldName in fieldNameList)
            {
                string result = _GetCell(sheet, j, fieldName, headerOrder);
                if (!string.IsNullOrEmpty(result))
                {
                    return result;
                }
            }
            return "";
        }

        private static bool _CheckNoDuplicateTestNamePass(Characterization item)
        {
            if (!item.Use.ToLower().Equals("use"))
            {
                return true;
            }
            /* check for duplicate tp name */
            if (AllcharList.ContainsKey(item.OriTpName.Replace(" ", "")))
            {
                const string outString = "Duplicate TP Name";
                Characterization myItem = AllcharList[item.OriTpName.Replace(" ", "")];
                ErrorManager.AddError(ErrorType.DuplicateTpName, myItem.SheetName, myItem.RowNum, myItem.ColNum("tpname(teusesuffix)"), myItem.Use, outString, myItem.OriTpName);
                ErrorManager.AddError(ErrorType.DuplicateTpName, item.SheetName, item.RowNum, myItem.ColNum("tpname(teusesuffix)"), item.Use, outString, item.TpName);

                foreach (int col in myItem.ColNum("tpname(teusesuffix)"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_DuplicateTpName_01, myItem.SheetName, myItem.RowNum, col, [],
                        new ErrorInfo() { Comments = new List<string>() { myItem.OriTpName } });
                    ErrorReportManager.AddError(CharErrorType.E_DuplicateTpName_01, item.SheetName, item.RowNum, col, [],
                        new ErrorInfo() { Comments = new List<string>() { item.TpName } });
                }
                return false;
            }

            AllcharList.Add(item.OriTpName.Replace(" ", ""), item);
            return true;
        }

        private static void _CheckTestNameFormula(ExcelWorksheet sheet, Characterization charRow)
        {
            if (charRow.SimTpName.ToLower().Trim('_') == charRow.OriTpName.ToLower().Trim('_'))
            {
                return;
            }

            const string outString = "Wrong equation of TP Name ";
            ErrorManager.AddError(ErrorType.WrongTpNameEquation, sheet.Name, charRow.RowNum, charRow.ColNum("tpname(teusesuffix)"), charRow.Use, outString);
            foreach (int col in charRow.ColNum("tpname(teusesuffix)"))
            {
                ErrorReportManager.AddError(CharErrorType.E_WrongTpNameEquation_01, sheet.Name, charRow.RowNum, col, []);
            }
        }

        private static bool _HasMergeLhv(Characterization charRow, IEnumerable<Characterization> charList)
        {
            /* search functional header tests according to Char_Input_Def, skip test or combine VDD data */

            foreach (Characterization item in charList.Where(item => _HasHlvCounterPart(item, charRow)))
            {
                // rename testName
                string newUserDef1 = item.UserDef1.Remove(item.UserDef1.Length - 1, 1) + "LH";
                item.TpName = item.TpName.Replace(item.UserDef1, newUserDef1);

                // update Use/Don't Use 
                if (Regex.IsMatch(charRow.Use, "^use", RegexOptions.IgnoreCase))
                {
                    item.Use = "Use";
                }

                if (_IsMarginHdbc(charRow.Category, item.Category))
                {
                    item.Category = item.Category.Substring(0, item.Category.Length - 2) + "HDBC";
                    item.TpName = Regex.Replace(item.TpName, "MarginHD|MarginBC", "MARGINHDBC",
                        RegexOptions.IgnoreCase);
                }
                return true;
            }
            return false;
        }

        private static bool _IsCheckHeaderFormatPass(string sheetName, int rowNum, Dictionary<string, int> headerOrder)
        {
            /*  Check missing Header and show in error report */
            var regexPowerRunScenarioWord = new Regex("^PowerRun(?<spell>.*)", RegexOptions.Singleline | RegexOptions.Compiled | RegexOptions.IgnoreCase);

            var essentialHeaderKeys = new List<string> {"use/notuse", "updated", "userdef1", "userdef2", "userdef3", "userdef4",
                "userdef5", "userdef6", "userdef7", "userdef8", "payload1", "payload2", "payload3", "payload4", "payload5",
                "group", "category", "method", "voltage", "usl", "lsl", "units", "init1", "init2", "init3", "init4",
                "init5", "init6", "init7", "init8", "init9", "init10", "priority", "testitems", "description"};

            var missingHeaderKeys = essentialHeaderKeys
                .Where(key => !headerOrder.Keys.Contains(key))
                .Select(key => "(" + key.ToUpper() + ")")
                .ToList();

            foreach (string headerKey in headerOrder.Keys)
            {
                if (regexPowerRunScenarioWord.IsMatch(headerKey))
                {
                    string spellCheck = regexPowerRunScenarioWord.Match(headerKey).Groups["spell"].ToString().ToLower();

                    if (spellCheck != "scenario")
                    {
                        missingHeaderKeys.Add("PowerRunScenario");
                        ErrorManager.AddError(ErrorType.WrongPowerRunScenario, sheetName, rowNum, 0, "Use", $"Spell wrong PowerRunScenario");
                        ErrorReportManager.AddError(CharErrorType.E_WrongPowerRunScenario_03, sheetName, rowNum, 0, []);
                    }
                }
            }

            if (missingHeaderKeys.Contains("(OTHERSUPPLIES)") && headerOrder.ContainsKey("voltagetable"))
            {
                missingHeaderKeys.Remove("(OTHERSUPPLIES)");
            }
            // report error
            if (missingHeaderKeys.Count <= 0)
            {
                return true;
            }

            string errorMessage = "Missing Header: " + string.Join(" ", missingHeaderKeys);
            ErrorManager.AddError(ErrorType.MissingHeader, sheetName, rowNum, 0, "Use", errorMessage);
            ErrorReportManager.AddError(CharErrorType.E_MissingHeader_01, sheetName, rowNum, 0, [string.Join(" ", missingHeaderKeys)]);
            MessageWriter.WriteMessage("Wrong header in input sheet " + sheetName + ", please check! ", EnumMessageLevel.Error);
            return false;
        }

        private static bool _HasHlvCounterPart(Characterization item, Characterization charRow)
        {
            // only check for functional user_def1
            if (!item.IsFuncRow())
            {
                return false;
            }

            if (!charRow.IsFuncRow())
            {
                return false;
            }

            // LVCC / HVCC should have same search algorithm
            if (charRow.Search != item.Search)
            {
                return false;
            }

            // bypassed LHV should not have search algorithm == "binary"
            if (charRow.Search.ToLower() == "binary")
            {
                return false;
            }

            // do not have counterpart for same user_def1
            if (charRow.UserDef1 == item.UserDef1)
            {
                return false;
            }

            // only have counterpart if user_def1 are only different at last character (MCL / MCH)
            if (charRow.UserDef1.Remove(charRow.UserDef1.Length - 1, 1) != item.UserDef1.Remove(item.UserDef1.Length - 1, 1))
            {
                return false;
            }

            string tpNameCheck = charRow.UserDef2 + "_" + charRow.UserDef3 + "_" + charRow.Group + "_" + charRow.UserDef4 + "_" + charRow.UserDef5 + "_" + charRow.UserDef6 + "_" + charRow.UserDef7 + "_" + charRow.UserDef8 + "_" + charRow.UserDef9;
            string mytpNameCheck = item.UserDef2 + "_" + item.UserDef3 + "_" + item.Group + "_" + item.UserDef4 + "_" + item.UserDef5 + "_" + item.UserDef6 + "_" + item.UserDef7 + "_" + item.UserDef8 + "_" + charRow.UserDef9;
            if (tpNameCheck != mytpNameCheck)
            {
                return false;
            }

            if (charRow.Category != item.Category && !_IsMarginHdbc(charRow.Category, item.Category))
            {
                return false;
            }

            return UtilityMain.UtilityFunction.CheckPowerCondition(charRow, item);
        }

        private static bool _IsMarginHdbc(string category1, string category2)
        {
            if (Regex.IsMatch(category1, "MarginHD", RegexOptions.IgnoreCase) && Regex.IsMatch(category2, "MarginBC", RegexOptions.IgnoreCase))
            {
                return true;
            }

            if (Regex.IsMatch(category2, "MarginHD", RegexOptions.IgnoreCase) && Regex.IsMatch(category1, "MarginBC", RegexOptions.IgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static void _UpdateExtraPatterns(ExcelWorksheet ws, int row, Characterization charItem, Dictionary<string, int> headerDictionary)
        {
            //todo need to sorting init and payload, init-1?
            List<string> keys = headerDictionary.Keys.ToList()
                .FindAll(p => Regex.IsMatch(p, @"init(?<Num>\d+)", RegexOptions.IgnoreCase));

            var initDic = new SortedDictionary<int, string>();

            foreach (string initKey in keys)
            {
                int.TryParse(Regex.Match(initKey, @"init(?<Num>\d+)", RegexOptions.IgnoreCase).Groups["Num"].ToString(), out int num);
                if (num <= 10)
                {
                    continue;
                }

                initDic.Add(num, initKey);
            }

            foreach (string init in initDic.Values.Where(init => _GetCell(ws, row, init, headerDictionary) != ""))
            {
                charItem.ExtraInits.Add(_GetCell(ws, row, init, headerDictionary));
            }

            keys = headerDictionary.Keys.ToList().FindAll(p => Regex.IsMatch(p, @"payload(?<Num>\d+)", RegexOptions.IgnoreCase));
            var plDic = new SortedDictionary<int, string>();

            foreach (string plKey in keys)
            {
                int.TryParse(Regex.Match(plKey, @"payload(?<Num>\d+)", RegexOptions.IgnoreCase).Groups["Num"].ToString(), out int num);
                if (num <= 5)
                {
                    continue;
                }

                plDic.Add(num, plKey);
            }

            foreach (string pl in plDic.Values.Where(pl => _GetCell(ws, row, pl, headerDictionary) != ""))
            {
                charItem.ExtraPLs.Add(_GetCell(ws, row, pl, headerDictionary));
            }
        }

        //20230725 take out for new mapping rule
        //private static void _UpdateDcAcCategory()
        //{
        //    foreach (var planSheet in CharPlanSheetDict.Where(planSheet => planSheet.Value.Count != 0))
        //    {
        //        foreach (var charRow in planSheet.Value)
        //        {
        //            InputDefRow inputDefRow;
        //            InputDefReader.BlockMapDict.TryGetValue(charRow.Category.Replace("HDBC", "HD") + ":" + charRow.Group, out inputDefRow);
        //            if (inputDefRow == null)
        //                InputDefReader.BlockMapDict.TryGetValue(charRow.Category.Replace("HDBC", "HD"), out inputDefRow);

        //            // ac category
        //            if (string.IsNullOrEmpty(charRow.AcCateName))
        //                charRow.AcCateName = inputDefRow != null ? inputDefRow.AcCategory : charRow.Category;
        //            charRow.AcCateNameOri = inputDefRow != null ? inputDefRow.AcCategory : charRow.Category;//for step-by-step ac generation

        //            // dc category
        //            if (charRow.OtherSupplies.Split(' ').Count() > 1)
        //            {
        //                charRow.DcCateName = charRow.OtherSupplies.Split(' ').First();
        //            }
        //            else
        //            {
        //                charRow.DcCateName = charRow.MappingSpec.DcCategory.Any() ? charRow.MappingSpec.DcCategory.First() : charRow.Category;
        //                //if (HardIpSheets.Contains(planSheet.Key.ToUpper())) continue;
        //                //if (charRow.Group.Contains("999")) continue;
        //                //if (inputDefRow != null && !inputDefRow.UsePMode) continue;

        //                //// edit dc category with performance mode
        //                //var tokens = charRow.DcCateName.Split('_');
        //                //if (tokens.Length < 4) // old style SOC_MS001
        //                //    charRow.DcCateName = charRow.DcCateName.PadLeft(3).Substring(0, 3) + "_" + charRow.Group;
        //                //else
        //                //{
        //                //    tokens[3] = charRow.Group;
        //                //    charRow.DcCateName = String.Join("_", tokens);
        //                //}
        //            }
        //        }
        //    }
        //}

        private static List<Characterization> _ExtendCharItemWithAC(Characterization item)
        {
            var result = new List<Characterization> { item };
            ShmooSpec stepAc = item.PowerSupplyX.FirstOrDefault(x => x.PowerCondition == "" && x.Type != "VDD"
                            && !string.IsNullOrEmpty(x.Start) && !string.IsNullOrEmpty(x.Stop));
            if (stepAc != null)
            {
                if (stepAc.Start != "" && stepAc.Stop != "")
                {
                    int indexAcItem = item.PowerSupplyX.IndexOf(stepAc);
                    try
                    {
                        result = new List<Characterization>();
                        string start = UtilityFunction.ConvertUnits(stepAc.Start);
                        string stop = UtilityFunction.ConvertUnits(stepAc.Stop);
                        if (stepAc.Start == stepAc.Stop)
                        {
                            var newitem = new Characterization(item);
                            double acValue = Convert.ToDouble(start);
                            var newAc = new ShmooSpec(stepAc)
                            {
                                Start = acValue.ToString(),
                                Stop = acValue.ToString(),
                                Step = "" //force AC value                            
                            };
                            newitem.PowerSupplyX[indexAcItem] = newAc;
                            //newitem.TpName = newitem.TpName.Replace(string.Format("_{0}_", newitem.UserDef7),
                            //    string.Format("_{0}{1}_", newitem.UserDef7, UtilityFunction.ConvertShiftSpeed(newAC.Start)));
                            newitem.AcStepName = item.Category + "_" + UtilityFunction.ConvertShiftSpeed(newAc.Start).Replace(".", "p") + "MHz";
                            if (stepAc.Type == "FRC")
                            {
                                newitem.FreeRunningClock = stepAc.Name + ":" + newAc.Start;
                            }

                            result.Add(newitem);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(stepAc.Step))
                            {
                                string step = UtilityFunction.ConvertUnits(stepAc.Step);
                                double points = (Convert.ToDouble(start) - Convert.ToDouble(stop)) / Convert.ToDouble(step);
                                for (int i = 0; i <= Math.Abs(points); i++)
                                {
                                    var newitem = new Characterization(item);
                                    double acValue = Convert.ToDouble(start) + (i * Convert.ToDouble(step));
                                    var newAc = new ShmooSpec(stepAc)
                                    {
                                        Start = acValue.ToString(),
                                        Stop = acValue.ToString(),
                                        Step = "" //force AC value                                    
                                    };
                                    newitem.PowerSupplyX[indexAcItem] = newAc;
                                    newitem.TpName = newitem.TpName.Replace($"_{newitem.UserDef7}_",
                                        $"_{newitem.UserDef7}{UtilityFunction.ConvertShiftSpeed(newAc.Start)}_");
                                    newitem.AcStepName = item.Category + "_" + UtilityFunction.ConvertShiftSpeed(newAc.Start).Replace(".", "p") + "MHz";
                                    if (stepAc.Type == "FRC")
                                    {
                                        newitem.FreeRunningClock = stepAc.Name + ":" + newAc.Start;
                                    }

                                    result.Add(newitem);
                                    //_AddAcByShmooStep(newitem.AcCateName, lRelatedUD7Var.Name, acValue);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            return result;
        }
        private static void _GetDigSrc(List<Characterization> charList)
        {
            foreach (Characterization charItem in charList)
            {
                var digSrcPatterns = new Dictionary<string, string>();
                var allPatterns = new List<string>
                {
                    charItem.Init1,
                    charItem.Init2,
                    charItem.Init3,
                    charItem.Init4,
                    charItem.Init5,
                    charItem.Init6,
                    charItem.Init7,
                    charItem.Init8,
                    charItem.Init9,
                    charItem.Init10,
                    charItem.Payload1,
                    charItem.Payload2,
                    charItem.Payload3,
                    charItem.Payload4,
                    charItem.Payload5,
                };

                _CheckDigSrcPat(charItem.Init1, "init1", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init2, "init2", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init3, "init3", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init4, "init4", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init5, "init5", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init6, "init6", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init7, "init7", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init8, "init8", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init9, "init9", digSrcPatterns);
                _CheckDigSrcPat(charItem.Init10, "init10", digSrcPatterns);

                string userdef6 = charItem.UserDef6;
                string userdef7 = charItem.UserDef7;
                string userdef8 = charItem.UserDef8;
                var srcPatterns = new Dictionary<int, string>();
                for (int i = 0; i < allPatterns.Count; i++)
                {
                    if (Regex.IsMatch(allPatterns[i], ":DigSrc", RegexOptions.IgnoreCase))
                    {
                        srcPatterns.Add(i, allPatterns[i]);
                    }
                }

                string[] array = new string[15];
                string sgmtdef = "sgmtdef" + Revision.DefaultSrcValue;
                switch (srcPatterns.Count)
                {
                    case 0:
                        break;
                    case 1:
                        int index = srcPatterns.Keys.ToList()[0];
                        string srcStr;
                        if (_CheckDigSrcValidData(userdef6, allPatterns[index], charItem, "userdef6"))
                        {
                            srcStr = userdef6;
                        }
                        else if (_CheckDigSrcValidData(userdef7, allPatterns[index], charItem, "userdef7"))
                        {
                            srcStr = userdef7;
                        }
                        else if (_CheckDigSrcValidData(userdef8, allPatterns[index], charItem, "userdef8")) //for proximaTC 2022/11/30
                        {
                            srcStr = userdef8;
                        }
                        else
                        {
                            ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(new List<string> { "userdef6", "userdef7", "userdef8" }),
                                charItem.Use, "There is one pattern specified with ':DigSrc' but no digsrc assignment in Userdef6/7/8");
                            foreach (int col in charItem.ColNum(new List<string> { "userdef6", "userdef7", "userdef8" }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_01, charItem.SheetName, charItem.RowNum, col, []);
                            }
                            break;
                        }
                        if (Regex.IsMatch(userdef6, "sgmt", RegexOptions.IgnoreCase))
                        {
                            srcStr = sgmtdef + srcStr;
                        }

                        array[index] = srcStr;
                        break;

                    case 2:
                        int index1 = srcPatterns.Keys.ToList()[0];
                        int index2 = srcPatterns.Keys.ToList()[1];
                        string srcStr1 = userdef6;
                        string srcStr2 = userdef7;

                        if (Regex.IsMatch(userdef6, "sgmt", RegexOptions.IgnoreCase))
                        {
                            srcStr1 = sgmtdef + srcStr1;
                        }

                        if (Regex.IsMatch(userdef7, "sgmt", RegexOptions.IgnoreCase))
                        {
                            srcStr2 = sgmtdef + srcStr2;
                        }

                        array[index1] = srcStr1;
                        array[index2] = srcStr2;
                        if (!_CheckDigSrcValidData(userdef6, allPatterns[index1], charItem, "userdef6") ||
                            !_CheckDigSrcValidData(userdef7, allPatterns[index2], charItem, "userdef7"))
                        {
                            ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(new List<string> { "userdef6", "userdef7" }),
                                charItem.Use, "There are two patterns specified with ':DigSrc' but missing digsrc assignment in Userdef6/7");
                            foreach (int col in charItem.ColNum(new List<string> { "userdef6", "userdef7" }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_02, charItem.SheetName, charItem.RowNum, col, []);
                            }
                        }

                        break;

                    default:
                        ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                            charItem.ColNum(digSrcPatterns.Keys.ToList())
                                , charItem.Use, "There are more than two patterns specified with ':DigSrc'");
                        foreach (int col in charItem.ColNum(digSrcPatterns.Keys.ToList()))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_03, charItem.SheetName, charItem.RowNum, col, []);
                        }
                        break;
                }

                _GetDigSrcReg(array, allPatterns);
                charItem.DigSrc =
                    array.All(string.IsNullOrEmpty) ?
                        "" : string.Join(",", array);
            }
        }
        private static void _GetDigSrcForNewTChar(List<Characterization> charList)
        {
            foreach (Characterization charItem in charList)
            {
                _GetDigSrcSRMforNewTChar(charItem.PatternCellList, charItem);
                _GetDigSrcRegforNewTChar(charItem.PatternCellList);
                _GetDigSrcEMAforNewTChar(charItem.PatternCellList, charItem);
                var result = new List<string>();
                var digSrcPatterns = new List<PatternCell>();
                foreach (PatternCell patternCell in charItem.PatternCellList)
                {
                    if (Regex.IsMatch(patternCell.PatternDefine, ":DigSrc", RegexOptions.IgnoreCase))
                    {
                        digSrcPatterns.Add(patternCell);
                    }
                }

                string userdef6 = charItem.UserDef6;
                string userdef7 = charItem.UserDef7;
                string userdef8 = charItem.UserDef8;

                string sgmtdef = "sgmtdef" + Revision.DefaultSrcValue;
                switch (digSrcPatterns.Count)
                {
                    case 0:
                        break;
                    case 1:
                        string srcStr;
                        if (_CheckDigSrcValidData(userdef6, digSrcPatterns[0].PatternDefine, charItem, "userdef6"))
                        {
                            srcStr = userdef6;
                        }
                        else if (_CheckDigSrcValidData(userdef7, digSrcPatterns[0].PatternDefine, charItem, "userdef7"))
                        {
                            srcStr = userdef7;
                        }
                        else if (_CheckDigSrcValidData(userdef8, digSrcPatterns[0].PatternDefine, charItem, "userdef8")) //for proximaTC 2022/11/30
                        {
                            srcStr = userdef8;
                        }
                        else
                        {
                            ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(new List<string> { "userdef6", "userdef7", "userdef8" }),
                                charItem.Use, "There is one pattern specified with ':DigSrc' but no digsrc assignment in Userdef6/7/8");
                            foreach (int col in charItem.ColNum(new List<string> { "userdef6", "userdef7", "userdef8" }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_01, charItem.SheetName, charItem.RowNum, col, []);
                            }
                            break;
                        }

                        _SetDigSrcArgs(digSrcPatterns[0], srcStr);
                        HandleSelsramInSgmtForNewTChar(charItem, digSrcPatterns[0], srcStr);
                        break;

                    case 2:
                        bool hasError = false;
                        if (!_CheckDigSrcValidData(userdef6, digSrcPatterns[0].PatternDefine, charItem, "userdef6"))
                        {
                            ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(new List<string> { "userdef6" }),
                                charItem.Use, "There are two patterns specified with ':DigSrc' but missing digsrc assignment in Userdef6");
                            foreach (int col in charItem.ColNum(new List<string> { "userdef6" }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_04, charItem.SheetName, charItem.RowNum, col, []);
                            }
                            hasError = true;
                        }

                        if (!_CheckDigSrcValidData(userdef7, digSrcPatterns[1].PatternDefine, charItem, "userdef7"))
                        {
                            ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(new List<string> { "userdef7" }),
                                charItem.Use, "There are two patterns specified with ':DigSrc' but missing digsrc assignment in Userdef7");
                            foreach (int col in charItem.ColNum(new List<string> { "userdef7" }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_05, charItem.SheetName, charItem.RowNum, col, []);
                            }
                            hasError = true;
                        }

                        if (hasError)
                        {
                            break;
                        }

                        string srcStr1 = userdef6;
                        string srcStr2 = userdef7;

                        if (Regex.IsMatch(userdef6, "sgmt", RegexOptions.IgnoreCase))
                        {
                            srcStr1 = sgmtdef + srcStr1;
                        }

                        if (Regex.IsMatch(userdef7, "sgmt", RegexOptions.IgnoreCase))
                        {
                            srcStr2 = sgmtdef + srcStr2;
                        }

                        _SetDigSrcArgs(digSrcPatterns[0], srcStr1);
                        HandleSelsramInSgmtForNewTChar(charItem, digSrcPatterns[0], srcStr1);
                        _SetDigSrcArgs(digSrcPatterns[1], srcStr2);
                        HandleSelsramInSgmtForNewTChar(charItem, digSrcPatterns[1], srcStr2);

                        break;

                    default:
                        ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                            charItem.ColNum(digSrcPatterns.Select(x => x.Header).ToList())
                                , charItem.Use, "There are more than two patterns specified with ':DigSrc'");
                        foreach (int col in charItem.ColNum(digSrcPatterns.Select(x => x.Header).ToList()))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_03, charItem.SheetName, charItem.RowNum, col, []);
                        }
                        break;
                }
                AppendZeroAsDefaultBit(charItem);
            }
        }

        private static void _GetDigSrcForMTDChar(List<Characterization> charList)
        {
            foreach (Characterization charItem in charList)
            {
                var result = new List<string>();
                var digSrcPatterns = new List<PatternCell>();
                foreach (PatternCell patternCell in charItem.PatternCellList)
                {
                    if (Regex.IsMatch(patternCell.PatternDefine, ":DigSrc", RegexOptions.IgnoreCase))
                    {
                        digSrcPatterns.Add(patternCell);
                    }
                }

                string userdef6 = charItem.UserDef6;
                string userdef7 = charItem.UserDef7;
                string userdef8 = charItem.UserDef8;

                string sgmtdef = "sgmtdef" + Revision.DefaultSrcValue;
                switch (digSrcPatterns.Count)
                {
                    case 0:
                        break;
                    case 1:
                        string srcStr = userdef6;


                        _SetDigSrcMTDArgs(digSrcPatterns[0], srcStr);

                        break;

                    default:
                        ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                            charItem.ColNum(digSrcPatterns.Select(x => x.Header).ToList())
                                , charItem.Use, "There are more than two MTD patterns specified with ':DigSrc'");
                        foreach (int col in charItem.ColNum(digSrcPatterns.Select(x => x.Header).ToList()))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_06, charItem.SheetName, charItem.RowNum, col, []);
                        }
                        break;
                }
                //AppendZeroAsDefaultBit(charItem);
            }
        }
        private static void AppendZeroAsDefaultBit(Characterization charItem)
        {
            foreach (PatternCell patternCell in charItem.PatternCellList)
            {
                if (string.IsNullOrEmpty(patternCell.DigSrcEq))
                {
                    continue;
                }

                string[] eqList = patternCell.DigSrcEq.Split('+');
                string[] segList = patternCell.DigSrcSeg.Split(';');
                var newSegList = new List<string>();
                foreach (string eq in eqList)
                {
                    string index = eq.Split('_')[0];
                    string length = eq.Split('_')[1];

                    string seg = segList.FirstOrDefault(x => x.Split('=')[0].Equals(index, StringComparison.OrdinalIgnoreCase));
                    if (seg != null)
                    {
                        newSegList.Add(seg);
                    }
                    else
                    {
                        newSegList.Add(string.Format($"{index}=0b{new string('0', int.Parse(length))}"));
                    }
                }
                patternCell.DigSrcSeg = string.Join(";", newSegList);
            }
        }
        private static void HandleSelsramInSgmtForNewTChar(Characterization charItem, PatternCell patternCell, string userdef)
        {
            int sTotalCount = charItem.UserDef9.ToUpper().Replace("SELSRAM", "").Replace("SELSRM", "").Count(x => x == 'S');
            int sRemindCount = sTotalCount;
            string[] splitSegments = patternCell.DigSrcSeg.Split(';');
            var newSegments = new List<string>();
            string[] splitEQs = patternCell.DigSrcEq.Split('+');
            var newEQs = new List<string>();
            var removeEq = new List<string>();
            var convertedEq = new List<string>();
            bool previousIsSelsram = false;
            string firstSegmentOfS = "";
            if (sTotalCount == 0)
            {
                return;
            }

            try
            {
                foreach (string segment in splitSegments)
                {
                    if (string.IsNullOrEmpty(segment))
                    {
                        continue;
                    }

                    string index = segment.Split('=')[0];
                    string value = segment.Split('=')[1];
                    string valueWithout0B = value.Replace("0b", "");
                    if (valueWithout0B.ToUpper().Contains("S"))
                    {
                        if (valueWithout0B.ToUpper().Replace("S", "") != "")
                        {
                            ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(userdef)
                                    , charItem.Use, string.Format($"S and binary can't be mixed in {index} ."));
                            foreach (int col in charItem.ColNum(userdef))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_07, charItem.SheetName, charItem.RowNum, col, [index]);
                            }
                            previousIsSelsram = false;
                        }
                        else if (!string.IsNullOrEmpty(valueWithout0B))
                        {
                            convertedEq.Add(index);
                            int sCount = valueWithout0B.Length;
                            string eq = splitEQs.FirstOrDefault(x => x.Split('_')[0].Equals(index, StringComparison.OrdinalIgnoreCase));
                            string bitLength = eq?.Split('_')[1] ?? "";

                            if (string.IsNullOrEmpty(eq))
                            {
                                ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(userdef)
                                    , charItem.Use, string.Format($"{index} doesn't exist in Send Bit Str of pattern info."));
                                foreach (int col in charItem.ColNum(userdef))
                                {
                                    ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_08, charItem.SheetName, charItem.RowNum, col, [index]);
                                }
                                previousIsSelsram = false;
                            }
                            else if (sRemindCount < sCount)
                            {
                                ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(userdef)
                                    , charItem.Use, string.Format($"{index} has S that exceed the length of {userdef}."));
                                foreach (int col in charItem.ColNum(userdef))
                                {
                                    ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_09, charItem.SheetName, charItem.RowNum, col, [index, userdef]);
                                }
                                previousIsSelsram = false;
                            }
                            else if (sCount != int.Parse(bitLength))
                            {
                                ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(userdef)
                                    , charItem.Use, string.Format($"S Length of {index} ({sCount}) isn't the same as segment length in pattern info ({bitLength})."));
                                foreach (int col in charItem.ColNum(userdef))
                                {
                                    ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_10, charItem.SheetName, charItem.RowNum, col, [index, $"{sCount}", bitLength]);
                                }
                                previousIsSelsram = false;
                            }
                            else if (sRemindCount < int.Parse(bitLength))
                            {
                                ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                charItem.ColNum(userdef)
                                    , charItem.Use, string.Format($"{string.Join(",", convertedEq)} over the S length of {charItem.UserDef9} ."));
                                foreach (int col in charItem.ColNum(userdef))
                                {
                                    ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_11, charItem.SheetName, charItem.RowNum, col, [string.Join(",", convertedEq)]);
                                }
                                previousIsSelsram = false;
                            }
                            else
                            {
                                if (sRemindCount == sTotalCount)//The first segment for S.
                                {
                                    value = "SELSRAM";
                                    previousIsSelsram = true;
                                    firstSegmentOfS = index;
                                    sRemindCount -= int.Parse(bitLength);
                                }
                                else if (previousIsSelsram)
                                {
                                    removeEq.Add(index);
                                    sRemindCount -= int.Parse(bitLength);
                                    continue;
                                }
                                else if (!previousIsSelsram)
                                {
                                    ErrorManager.AddError(ErrorType.WrongDigSrc, charItem.SheetName, charItem.RowNum,
                                    charItem.ColNum(userdef)
                                        , charItem.Use, string.Format($"{string.Join(",", convertedEq)}: segments for S are not continuous."));
                                    foreach (int col in charItem.ColNum(userdef))
                                    {
                                        ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_12, charItem.SheetName, charItem.RowNum, col, [string.Join(",", convertedEq)]);
                                    }
                                    return;
                                }
                            }
                        }
                    }
                    else
                    {
                        previousIsSelsram = false;
                    }
                    newSegments.Add(string.Format($"{index}={value}"));
                }
                newEQs = RebuildDigSrcEqs(splitEQs, firstSegmentOfS, removeEq, sTotalCount);
                patternCell.DigSrcSeg = string.Join(";", newSegments);
                patternCell.DigSrcEq = string.Join("+", newEQs);
            }
            catch (Exception)
            {
            }

        }

        private static List<string> RebuildDigSrcEqs(string[] splitEQs, string firstSegmentOfS, List<string> removeEq, int sTotalCount)
        {
            var newEQs = new List<string>();
            foreach (string eq in splitEQs)
            {
                if (string.IsNullOrEmpty(eq))
                {
                    continue;
                }

                string index = eq.Split('_')[0];
                string length = eq.Split('_')[1];

                if (firstSegmentOfS.Equals(index, StringComparison.OrdinalIgnoreCase))
                {
                    newEQs.Add(string.Format($"{index}_{sTotalCount}"));
                }
                else if (!removeEq.Any(x => x.Equals(index, StringComparison.OrdinalIgnoreCase)))
                {
                    newEQs.Add(string.Format($"{index}_{length}"));
                }
            }
            return newEQs;
        }
        private static void _SetDigSrcArgs(PatternCell patternCell, string digSrcStr)
        {
            string pattern = patternCell.PatternDefine;
            string patternName = pattern.Split(':')[0].ToUpper();
            if (!UtilityMain.UtilityData.PatInfoDict.TryGetValue(patternName, out HardIpReference patternInfo))
            {
                return;
            }

            if (string.IsNullOrEmpty(patternInfo.SendBitStr))
            {
                return;
            }

            patternCell.DigSrcBitSize = patternInfo.SendBit.ToString();
            patternCell.DigSrcEq = patternInfo.SendBitStr;
            patternCell.DigSrcPin = patternInfo.SendPinName;

            var digSrcSegFromStr = new List<string>();
            if (Regex.IsMatch(digSrcStr, "sgmt", RegexOptions.IgnoreCase) || Regex.IsMatch(digSrcStr, "^[01]+$"))
            {
                string dataContent = Regex.Replace(digSrcStr, "^sgmtdef", "", RegexOptions.IgnoreCase);
                string[] sgmtSets = Regex.Split(dataContent, @"(sgmt\d+)", RegexOptions.IgnoreCase);
                var sgmtContent = new Dictionary<string, string>();

                List<string> sendBitStrList = patternInfo.SendBitStr.Split('+').ToList();
                List<string> sendBitNameList = patternInfo.SendBitName.Split('+').ToList();
                if (sgmtSets.Length > 1)
                {
                    for (int i = 0; i < sgmtSets.Length; i++)
                    {
                        string sgmtName = Regex.Match(sgmtSets[i], @"sgmt\d+", RegexOptions.IgnoreCase).ToString().ToLower();

                        if (sgmtName == "" || i + 1 >= sgmtSets.Length)
                        {
                            continue;
                        }

                        if (Regex.IsMatch(sgmtSets[i + 1], "^f", RegexOptions.IgnoreCase))
                        {
                            digSrcSegFromStr.Add(
                                $"{sgmtSets[i]}={"0b" + new string(sgmtSets[i + 1].Skip(1).ToArray())}");
                        }
                        else if (Regex.IsMatch(sgmtSets[i + 1], "^g", RegexOptions.IgnoreCase))
                        {
                            digSrcSegFromStr.Add(
                                $"{sgmtSets[i]}={"0x" + new string(sgmtSets[i + 1].Skip(1).ToArray())}");
                        }
                    }
                }
                else
                {
                    string sourctBits = sgmtSets[0];
                    string sgmt0 = sendBitStrList.FirstOrDefault(x => x.ToLower().Contains("sgmt0"));
                    if (!string.IsNullOrEmpty(sgmt0))
                    {
                        if (sgmt0.Split('_').Length >= 2)
                        {
                            if (int.TryParse(sgmt0.Split('_')[1], out int sgmt0Length))
                            {

                                if (sgmtSets[0].Length > sgmt0Length)
                                {
                                    sourctBits = sourctBits.Remove(0, sgmtSets[0].Length - sgmt0Length);
                                }
                                else
                                {
                                    sourctBits = sourctBits.PadLeft(sgmt0Length, '0');
                                }
                            }
                        }
                    }
                    digSrcSegFromStr.Add($"{"sgmt0"}={"0b" + sourctBits}");
                }
                patternCell.DigSrcSeg = string.Join(";", digSrcSegFromStr);
            }
        }

        private static void _SetDigSrcMTDArgs(PatternCell patternCell, string digSrcStr)
        {
            var sendBitConvert = new List<string>();
            string pattern = patternCell.PatternDefine;
            string patternName = pattern.Split(':')[0].ToUpper();
            if (!UtilityMain.UtilityData.PatInfoDict.TryGetValue(patternName, out HardIpReference patternInfo))
            {
                return;
            }

            if (string.IsNullOrEmpty(patternInfo.SendBitStr))
            {
                return;
            }

            if (string.IsNullOrEmpty(patternInfo.SendBitName))
            {
                return;
            }

            string sendBintCountStr = patternInfo.SendBitStr.Split('+').First().Split('_').Last();
            int.TryParse(sendBintCountStr, out int sendBintCount);

            var digSrcApkDictionary = new Dictionary<string, string>();

            foreach (string digSrcApkList in digSrcStr.Split(';'))
            {

                string[] digSrcApkData = digSrcApkList.Split(':');
                string digSrcApk = digSrcApkData.First().ToUpper();
                string digSrcApkValue = digSrcApkData.Last().Replace("0x", "");
                int tempHexValue = Convert.ToInt32(digSrcApkValue, 16);
                string tempBinValue = Convert.ToString(tempHexValue, 2).PadLeft(sendBintCount, '0');
                if (!digSrcApkDictionary.ContainsKey(digSrcApk))
                {
                    digSrcApkDictionary.Add(digSrcApk, tempBinValue);
                }
            }


            var sendBitNameList = patternInfo.SendBitName.Split('+').Select(p => p.Split(new string[] { "__" }, StringSplitOptions.None)[0]).GroupBy(x => x).Select(g => new { Key = g.Key, Count = g.Count() });

            foreach (var sendBitApk in sendBitNameList)
            {
                string sendBitApkName = sendBitApk.Key.ToUpper();
                bool sendValue = digSrcApkDictionary.TryGetValue(sendBitApkName, out string actualValue);
                string sendBitTempStr = actualValue != null ? $"{actualValue}@{sendBitApk.Count}" : Convert.ToString(0, 2).PadLeft(sendBintCount, '0');
                sendBitConvert.Add(sendBitTempStr);

            }
            patternCell.DigSrcSeg = string.Join("+", sendBitConvert);
            patternCell.DigSrcPin = patternInfo.SendPinName;

        }

        private static void _GetDigSrcReg(string[] digSrcArr, List<string> patterns)
        {
            if (DigSrcReg == null)
            {
                return;
            }

            for (int p = 0; p < patterns.Count; p++)
            {
                string pattern = patterns[p];
                string patternName = pattern.Split(':')[0].ToUpper();
                string label = pattern.Split(':').Length > 1 ? pattern.Split(':')[1].Trim().ToUpper() : "";
                if (string.IsNullOrEmpty(label) ||
                    !UtilityMain.UtilityData.PatInfoDict.ContainsKey(patternName) ||
                    !DigSrcReg.Keys.Any(x => x.Equals(label, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                List<DigSrcRegData> digSrcValueList = DigSrcReg.FirstOrDefault(x => x.Key.Equals(label, StringComparison.OrdinalIgnoreCase)).Value;
                HardIpReference patteanInfo = UtilityMain.UtilityData.PatInfoDict[patternName];
                List<string> sendBitStrList = patteanInfo.SendBitStr.Split('+').ToList();
                List<string> sendBitNameList = patteanInfo.SendBitName.Split('+').ToList();

                if (sendBitStrList.Count != sendBitNameList.Count)
                {
                    continue;
                }

                var sendBitMap = new Dictionary<string, List<string>>();
                var convertStrList = new List<string>();
                for (int i = 0; i < sendBitStrList.Count; i++)
                {
                    string sendBitStr = sendBitStrList[i];
                    string sendBitName = sendBitNameList[i];
                    DigSrcRegData digSrcValue = digSrcValueList.FirstOrDefault(x => x.RegName.Equals(sendBitName, StringComparison.OrdinalIgnoreCase));
                    if (digSrcValue != null && sendBitStr.Split('_').Length == 2)
                    {
                        string sgmtName = sendBitStr.Split('_')[0];
                        if (!int.TryParse(sendBitStr.Split('_')[1], out int sgmtLength))
                        {
                            continue;
                        }

                        string convertStr = _ConvertToSgmt(sgmtName, digSrcValue.RegValue);
                        if (!string.IsNullOrEmpty(convertStr))
                        {
                            convertStrList.Add(convertStr);
                        }
                    }
                }
                digSrcArr[p] = string.Join("", convertStrList);
                patterns[p] = patternName + ":DigSrc";
            }
        }

        private static void _GetDigSrcRegforNewTChar(List<PatternCell> patternCells)
        {
            if (DigSrcReg == null)
            {
                return;
            }

            foreach (PatternCell patternCell in patternCells)
            {
                string patternName = patternCell.PatternDefine.Split(':')[0].ToUpper();
                string label = patternCell.PatternDefine.Split(':').Length > 1 ? patternCell.PatternDefine.Split(':')[1].Trim().ToUpper() : "";
                if (string.IsNullOrEmpty(label) ||
                    !UtilityMain.UtilityData.PatInfoDict.ContainsKey(patternName) ||
                    !DigSrcReg.Keys.Any(x => x.Equals(label, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                List<DigSrcRegData> digSrcValueList = DigSrcReg.FirstOrDefault(x => x.Key.Equals(label, StringComparison.OrdinalIgnoreCase)).Value;

                HardIpReference patternInfo = UtilityMain.UtilityData.PatInfoDict[patternName];
                if (string.IsNullOrEmpty(patternInfo.SendBitStr))
                {
                    continue;
                }

                patternCell.DigSrcBitSize = patternInfo.SendBit.ToString();
                patternCell.DigSrcEq = patternInfo.SendBitStr;
                patternCell.DigSrcPin = patternInfo.SendPinName;

                List<string> sendBitStrList = patternInfo.SendBitStr.Split('+').ToList();
                List<string> sendBitNameList = patternInfo.SendBitName.Split('+').ToList();
                if (sendBitStrList.Count != sendBitNameList.Count)
                {
                    continue;
                }

                if (string.IsNullOrEmpty(patternCell.DigSrcSeg))
                {
                    var sendBitMap = new Dictionary<string, List<string>>();
                    var digSrcEqFromStr = new List<string>();
                    for (int i = 0; i < sendBitStrList.Count; i++)
                    {
                        string sendBitStr = sendBitStrList[i];
                        string sendBitName = sendBitNameList[i];
                        DigSrcRegData digSrcValue = digSrcValueList.FirstOrDefault(x => x.RegName.Equals(sendBitName, StringComparison.OrdinalIgnoreCase));
                        if (digSrcValue != null && sendBitStr.Split('_').Length == 2)
                        {
                            string sgmtName = sendBitStr.Split('_')[0];
                            if (!int.TryParse(sendBitStr.Split('_')[1], out int sgmtLength))
                            {
                                continue;
                            }

                            digSrcEqFromStr.Add($"{sgmtName}={digSrcValue.RegValue}");
                        }
                    }
                    patternCell.DigSrcSeg = string.Join(";", digSrcEqFromStr);
                }
            }
        }

        private static void _GetDigSrcSRMforNewTChar(List<PatternCell> patternCells, Characterization charItem)
        {
            foreach (PatternCell patternCell in patternCells)
            {
                string patternName = patternCell.PatternDefine.Split(':')[0].ToUpper();
                if (!patternName.Contains("SRMDSSC"))
                {
                    continue;
                }

                if (!UtilityMain.UtilityData.PatInfoDict.TryGetValue(patternName, out HardIpReference patternInfo))
                {
                    //_GetDigSrcSRMPlanforNewTChar(patternCell, charItem);
                    continue;
                }

                if (string.IsNullOrEmpty(patternInfo.SendBitStr))
                {
                    //_GetDigSrcSRMPlanforNewTChar(patternCell, charItem);
                    continue;
                }

                patternCell.DigSrcBitSize = patternInfo.SendBit.ToString();
                //patternCell.DigSrcEQ = patternInfo.SendBitStr;
                patternCell.DigSrcEq = "sgmt0_" + patternCell.DigSrcBitSize;
                patternCell.DigSrcPin = patternInfo.SendPinName;

                //var sendBitStrList = patternInfo.SendBitStr.Split('+').ToList();
                //var sendBitNameList = patternInfo.SendBitName.Split('+').ToList();
                //if (sendBitStrList.Count != sendBitNameList.Count)
                //    continue;
                //if (string.IsNullOrEmpty(patternCell.DigSrcSeg))
                //{
                //    var digSrcEqFromStr = new List<string>();
                //    for (int i = 0; i < sendBitStrList.Count; i++)
                //    {
                //        var sendBitStr = sendBitStrList[i];
                //        var digSrcValue = "SELSRAM";
                //        if (digSrcValue != null && sendBitStr.Split('_').Length == 2)
                //        {
                //            var sgmtName = sendBitStr.Split('_')[0];
                //            int sgmtLength;
                //            if (!int.TryParse(sendBitStr.Split('_')[1], out sgmtLength))
                //                continue;
                //            digSrcEqFromStr.Add(string.Format("{0}={1}", sgmtName, "SELSRAM"));
                //        }
                //    }
                //    patternCell.DigSrcSeg = string.Join(";", digSrcEqFromStr);
                //}
                patternCell.DigSrcSeg = "sgmt0=SELSRAM";
            }
        }

        private static void _GetDigSrcEMAforNewTChar(List<PatternCell> patternCells, Characterization charItem)
        {
            if (!UtilityMain.UtilityData.EmaMappingItems.Any())
            {
                return;
            }

            int userdefSmgtIdx = 0;
            foreach (PatternCell patternCell in patternCells)
            {
                string patternName = patternCell.PatternDefine.Split(':')[0].ToUpper();
                string label = patternCell.PatternDefine.Split(':').Length > 1 ? patternCell.PatternDefine.Split(':')[1].Trim().ToUpper() : "";
                if (!label.Equals("DIGSRC") ||
                    !UtilityMain.UtilityData.PatInfoDict.ContainsKey(patternName) ||
                    !UtilityMain.UtilityData.EmaMappingItems.Any(x => x.Pattern.Equals(patternName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                //Decide which userdef will be used
                userdefSmgtIdx++;
                EmaMappingItem emaByPattern = UtilityMain.UtilityData.EmaMappingItems.FirstOrDefault(x => x.Pattern.Equals(patternName, StringComparison.OrdinalIgnoreCase));
                if (emaByPattern == null)
                {
                    continue;
                }

                string userdef = "";
                while (userdefSmgtIdx <= 3)
                {
                    if (userdefSmgtIdx == 1)
                    {
                        if (emaByPattern.CasesList.Any(x => x.Equals(charItem.UserDef6, StringComparison.OrdinalIgnoreCase)))
                        {
                            userdef = charItem.UserDef6;
                            userdefSmgtIdx++;
                            break;
                        }
                    }
                    else if (userdefSmgtIdx == 2)
                    {
                        if (emaByPattern.CasesList.Any(x => x.Equals(charItem.UserDef7, StringComparison.OrdinalIgnoreCase)))
                        {
                            userdef = charItem.UserDef7;
                            userdefSmgtIdx++;
                            break;
                        }
                    }
                    else if (userdefSmgtIdx == 3)
                    {
                        if (emaByPattern.CasesList.Any(x => x.Equals(charItem.UserDef8, StringComparison.OrdinalIgnoreCase)))
                        {
                            userdef = charItem.UserDef8;
                            userdefSmgtIdx++;
                            break;
                        }
                    }
                    userdefSmgtIdx++;
                }
                if (string.IsNullOrEmpty(userdef))
                {
                    continue;
                }

                string emaByCase = emaByPattern.GetCaseData(userdef);

                HardIpReference patternInfo = UtilityMain.UtilityData.PatInfoDict[patternName];
                if (string.IsNullOrEmpty(patternInfo.SendBitStr))
                {
                    continue;
                }

                patternCell.DigSrcBitSize = patternInfo.SendBit.ToString();
                patternCell.DigSrcEq = patternInfo.SendBitStr;
                patternCell.DigSrcPin = patternInfo.SendPinName;

                List<string> sendBitStrList = patternInfo.SendBitStr.Split('+').ToList();

                var digSrcEqFromStr = new List<string>();
                if (Regex.IsMatch(emaByCase, "sgmt", RegexOptions.IgnoreCase) || Regex.IsMatch(emaByCase, "^[01]+$"))
                {
                    string dataContent = Regex.Replace(emaByCase, "^sgmtdef", "", RegexOptions.IgnoreCase);
                    string[] sgmtSets = Regex.Split(dataContent, @"(sgmt\d+)", RegexOptions.IgnoreCase);
                    var sgmtContent = new Dictionary<string, string>();

                    for (int i = 0; i < sgmtSets.Length; i++)
                    {
                        string sgmtName = Regex.Match(sgmtSets[i], @"sgmt\d+", RegexOptions.IgnoreCase).ToString().ToLower();

                        if (sgmtName == "" || i + 1 >= sgmtSets.Length)
                        {
                            continue;
                        }

                        if (Regex.IsMatch(sgmtSets[i + 1], "^f", RegexOptions.IgnoreCase))
                        {
                            digSrcEqFromStr.Add($"{sgmtSets[i]}={"0b" + new string(sgmtSets[i + 1].Skip(1).ToArray())}");
                        }
                        else if (Regex.IsMatch(sgmtSets[i + 1], "^g", RegexOptions.IgnoreCase))
                        {
                            digSrcEqFromStr.Add($"{sgmtSets[i]}={"0x" + new string(sgmtSets[i + 1].Skip(1).ToArray())}");
                        }
                    }
                }
                patternCell.DigSrcSeg = string.Join(";", digSrcEqFromStr);
            }
        }

        private static string _ConvertToSgmt(string sgmtName, string digSrcValue)
        {
            string sgmtStr = "";
            //Hex
            if (digSrcValue.StartsWith("0x") && digSrcValue.Length > 2)
            {
                sgmtStr = digSrcValue.Remove(0, 2);

                if (uint.TryParse(sgmtStr, System.Globalization.NumberStyles.HexNumber, null, out _))
                {
                    sgmtStr = sgmtName + "g" + sgmtStr;
                }
            }
            //Binary
            else if (Regex.IsMatch(digSrcValue, "^[01]+$"))
            {
                sgmtStr = sgmtName + "f" + digSrcValue;
            }
            return sgmtStr;
        }

        private static void _CheckDigSrcPat(string pattern, string key, IDictionary<string, string> dicSrcGroup)
        {
            if (Regex.IsMatch(pattern, ":DigSrc", RegexOptions.IgnoreCase))
            {
                dicSrcGroup.Add(key, pattern);
            }
        }

        private static bool _CheckDigSrcValidData(string data, string pattern, Characterization charItem, string header)
        {
            //if (UtilityMain.UtilityData.SourceMappingItems.Exists(p => p.Label.Equals(data, StringComparison.OrdinalIgnoreCase)))
            //return true;            
            if (header.ToUpper() == "USERDEF8")
            {
                if (Regex.IsMatch(data, "^[0-1]+$"))
                {
                    return true;
                }
            }
            if (UtilityMain.UtilityData.EmaMappingItems.Exists(p => p.Pattern.Equals(pattern.Split(':')[0],
                StringComparison.OrdinalIgnoreCase)))
            {
                EmaMappingItem emaItem =
                    UtilityMain.UtilityData.EmaMappingItems.FirstOrDefault(
                        p => p.Pattern.Equals(pattern.Split(':')[0], StringComparison.OrdinalIgnoreCase));
                if (emaItem.CasesList.Any(p => p.Equals(data, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            if (Regex.IsMatch(data, "sgmt", RegexOptions.IgnoreCase) || Regex.IsMatch(data, "^[01]+$"))
            {
                string dataContent = Regex.Replace(data, "^sgmtdef", "", RegexOptions.IgnoreCase);
                string[] sgmtSets = Regex.Split(dataContent, @"(sgmt\d+)", RegexOptions.IgnoreCase);
                var sgmtContent = new Dictionary<string, string>();

                for (int i = 0; i < sgmtSets.Length; i++)
                {
                    string sgmtName = Regex.Match(sgmtSets[i], @"sgmt\d+", RegexOptions.IgnoreCase).ToString().ToLower();

                    if (sgmtName == "" || i + 1 >= sgmtSets.Length)
                    {
                        continue;
                    }

                    string sgmtStr = "";
                    if (Regex.IsMatch(sgmtSets[i + 1], "^f", RegexOptions.IgnoreCase))
                    {
                        if (Regex.IsMatch(sgmtSets[i + 1], "^f.*[^01S]+", RegexOptions.IgnoreCase))
                        {
                            ErrorManager.Add(new ErrorMessage
                            {
                                ErrorLevel = ErrorLevel.Error,
                                ErrorType = ErrorType.WrongDigSrc,
                                SheetName = charItem.SheetName,
                                RowNum = charItem.RowNum,
                                ColList = charItem.ColNum(new List<string> { header }),
                                Message = $"Data in {sgmtSets[i] + sgmtSets[i + 1]} must be 0, 1, S",
                            });
                            foreach (int col in charItem.ColNum(new List<string> { header }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_13, charItem.SheetName, charItem.RowNum, col, [sgmtSets[i] + sgmtSets[i + 1]]);
                            }
                        }
                    }
                    else if (Regex.IsMatch(sgmtSets[i + 1], "^g", RegexOptions.IgnoreCase))
                    {
                        if (Regex.IsMatch(sgmtSets[i + 1], "^g.*[^0-9a-fA-F]+", RegexOptions.IgnoreCase))
                        {
                            ErrorManager.Add(new ErrorMessage
                            {
                                ErrorLevel = ErrorLevel.Error,
                                ErrorType = ErrorType.WrongDigSrc,
                                SheetName = charItem.SheetName,
                                RowNum = charItem.RowNum,
                                ColList = charItem.ColNum(new List<string> { header }),
                                Message = $"Data in {sgmtSets[i] + sgmtSets[i + 1]} must be 0-9 and A-F",
                            });
                            foreach (int col in charItem.ColNum(new List<string> { header }))
                            {
                                ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_14, charItem.SheetName, charItem.RowNum, col, [sgmtSets[i] + sgmtSets[i + 1]]);
                            }
                        }
                    }
                    else
                    {
                        ErrorManager.Add(new ErrorMessage
                        {
                            ErrorLevel = ErrorLevel.Error,
                            ErrorType = ErrorType.WrongDigSrc,
                            SheetName = charItem.SheetName,
                            RowNum = charItem.RowNum,
                            ColList = charItem.ColNum(new List<string> { header }),
                            Message = $"\"{sgmtSets[i] + sgmtSets[i + 1]}\" didn't be recognized",
                        });
                        foreach (int col in charItem.ColNum(new List<string> { header }))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongDigSrc_15, charItem.SheetName, charItem.RowNum, col, [sgmtSets[i] + sgmtSets[i + 1]]);
                        }
                    }
                    sgmtContent.Add(sgmtName, sgmtStr);
                }
                if (!UtilityMain.UtilityData.PatInfoDict.ContainsKey(pattern.Split(':')[0].ToUpper()))
                {
                    //Check pattern in pattern info is in DigSrcChecker.cs
                    return false;
                }
                HardIpReference patternInfo = UtilityMain.UtilityData.PatInfoDict[pattern.Split(':')[0].ToUpper()];
                List<string> sendBitStrList = patternInfo.SendBitStr.Split('+').ToList();
                var sendBitStrLengthDict = new Dictionary<string, int>();
                if (!string.IsNullOrEmpty(patternInfo.SendBitStr))
                {
                    foreach (string sgmt in sendBitStrList)
                    {
                        try
                        {
                            string sgmtName = sgmt.Split('_')[0].ToLower();
                            string sgmtLength = sgmt.Split('_')[1];
                            int bitCount = Convert.ToInt32(sgmtLength);
                            sendBitStrLengthDict[sgmtName] = bitCount;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    foreach (string sgmtDefined in sgmtContent.Keys)
                    {
                        if (!sendBitStrLengthDict.Keys.Contains(sgmtDefined, StringComparer.OrdinalIgnoreCase))
                        {
                            ErrorManager.Add(new ErrorMessage
                            {
                                ErrorLevel = ErrorLevel.Warning,
                                ErrorType = ErrorType.MissingDigSrcSgmtInPattern,
                                SheetName = charItem.SheetName,
                                RowNum = charItem.RowNum,
                                ColList = charItem.ColNum(new List<string> { header }),
                                Message = $"Pattern: \"{pattern.Split(':')[0]}\" has not \"{sgmtDefined}\""
                            });
                            foreach (int col in charItem.ColNum(new List<string> { header }))
                            {
                                ErrorReportManager.AddError(CharErrorType.W_MissingDigSrcSgmtInPattern_01,
                                    charItem.SheetName, charItem.RowNum, col, [pattern.Split(':')[0], sgmtDefined]);
                            }
                        }
                    }
                }
                return true;
            }
            return false;
        }

        protected class PinVarHeader
        {
            public int PinVar = -1;
            public int Start = -1;
            public int Stop = -1;
            public int Step = -1;
            public int Type = -1;
            public int Order = -1;
            public bool IsVar = false;
        }
    }
}
