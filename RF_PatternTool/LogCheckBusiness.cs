using System.Text;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.ErrorReport.Base;

using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using RF_PatternTool.PatternGen;
using RF_PatternTool.PatternStruct;

using RfLib.Dvdc.GenTemplate.Bussiness;

namespace RF_PatternTool
{
    public class LogCheckBusiness
    {
        private readonly BodyPattern _bodyP = new BodyPattern();
        private readonly ReadPattern _readP = new ReadPattern(0.0);
        private readonly ReadPattern _read64P = new ReadPattern(0.0);
        private readonly SubRoutine _subrP = new SubRoutine();
        private readonly WritePattern _writeP = new WritePattern(0.0);
        private readonly WritePattern _write64P = new WritePattern(0.0);
        private readonly MatchLoopPattern _matchP = new MatchLoopPattern(0.0);
        private readonly Dictionary<string, List<BenchLogItem>> _rEuseDictionary = new Dictionary<string, List<BenchLogItem>>();
        public List<BenchLogFile> LogFile = new List<BenchLogFile>();
        public string LogName;
        public List<List<string>> EfusebitDefRows = new List<List<string>>();
        public HashSet<string> LogTestName = new HashSet<string>();

        public HashSet<string> LogStoreName = new HashSet<string>();
        public HashSet<string> LogCalcStoreName = new HashSet<string>();
        public HashSet<string> LogMeasStoreName = new HashSet<string>();
        public HashSet<string> LogFullStoreName = new HashSet<string>();

        public PinMapSheet Pinmap;
        public List<string> PatNames = new List<string>();
        private readonly Dictionary<string, string> _srcrecords = new Dictionary<string, string>();
        private readonly Dictionary<string, List<string>> _srcwarnrecords = new Dictionary<string, List<string>>();
        private bool _isFW = false;
        private bool _isOverWrite = false;
        public int PatBit = 32;

        public Dictionary<string, List<string>> FunctionMaps = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        private void InitSetup()
        {
            _bodyP.ReadTemplate();
            _readP.ReadTemplate(Template.TemplateSet.ReadTemp);
            _read64P.ReadTemplate(Template.TemplateSet.Read64Temp);
            _subrP.ReadTemplate();
            _writeP.ReadTemplate(Template.TemplateSet.WriteTemp);
            _write64P.ReadTemplate(Template.TemplateSet.Write64Temp);
            _matchP.ReadTemplate();

            if (_readP.SourceCaptureDictionary["Source"].Count() > 32 ||
                _readP.SourceCaptureDictionary["Capture"].Count() > 32 ||
                _writeP.SourceCaptureDictionary["Source"].Count() / 2 > 32)
            {
                PatBit = 64;
                _readP.PatBit = 64;
                _read64P.PatBit = 64;
                _writeP.PatBit = 64;
                _write64P.PatBit = 64;
                _matchP.PatBit = 64;
            }
        }

        public static MappingItem GetNamingInfo(string filename)
        {
            var patternName = new MappingItem();
            string datecode = DateTime.Now.ToString("yyyyMMdd");
            BenchLogFile shortinfo = GetLogPatternName(filename);
            string logname = string.IsNullOrEmpty(shortinfo.PatternSubName)
                ? Path.GetFileNameWithoutExtension(filename)
                : shortinfo.PatternSubName;

            if (logname.Contains('#'))
            {
                patternName.Version = logname.Split('#')[1];
            }
            else
            {
                patternName.Version = "1";
            }

            patternName.Log = logname.Split('#')[0];
            patternName.Pattern = "" + logname.Split('#')[0];

            patternName.Silicon = "A0";
            patternName.Date = string.IsNullOrEmpty(shortinfo.DateCode) ? datecode : shortinfo.DateCode;
            patternName.Version = shortinfo.Version;
            patternName.Inits.AddRange(shortinfo.Inits);
            return patternName;

        }

        public KeyValuePair<string, List<string>> CheckBench(string filepath, bool isOverWrite, Dictionary<string, Register> tmpregisters, string outputDir)
        {
            var result = new KeyValuePair<string, List<string>>(outputDir, new List<string> { });

            string logDirect = Path.GetDirectoryName(filepath);
            string logname = Path.GetFileNameWithoutExtension(filepath); //todo
            LogName = logname;
            var registers = new Dictionary<string, Register>(StringComparer.OrdinalIgnoreCase);

            //Read Log
            LogFile = ReadLogFile(filepath, true);

            _isOverWrite = isOverWrite;

            _srcrecords.Clear();
            _srcwarnrecords.Clear();

            InitSetup();

            foreach (BenchLogFile log in LogFile)
            {
                MappingItem patternName = GetNamingInfo(filepath);

                string toolAtp = patternName.GetFullPattern() + ".atp";
                result.Value.Add(toolAtp);

                List<BenchLogItem> trimLogs = new List<BenchLogItem>();
                List<BenchLogItem> usedTrimLogs = new List<BenchLogItem>();
                #region Check Bench log
                var trimInfoList = new List<string> { "trim_start", "trim_end", "write_trim" };
                _isFW = log.Type.Equals("FARF", StringComparison.OrdinalIgnoreCase) ||
                    log.Logs.Exists(p => Regex.IsMatch(p.Operation, "Wait_For", RegexOptions.IgnoreCase));

                string patSubName = log.PatternSubName;
                _rEuseDictionary.Clear();

                BenchLogItem prevloginfo = null;

                foreach (BenchLogItem loginfo in log.Logs)
                {
                    try
                    {
                        ProcessBenchLogOperation(loginfo, logDirect, tmpregisters, patternName, registers, trimLogs, usedTrimLogs, ref trimInfoList, ref patSubName);

                        prevloginfo = loginfo;
                    }
                    catch (Exception)
                    {
                        //;
                    }
                }

                foreach (KeyValuePair<string, List<BenchLogItem>> efuse in _rEuseDictionary)
                {
                    if (string.IsNullOrEmpty(efuse.Key))
                    {
                        continue;
                    }

                    var fuseItems = efuse.Value.Select(p => p.GetAddressInfo()).Distinct().ToList();
                    if (fuseItems.Count() > 1)
                    {
                        foreach (BenchLogItem fuseitem in efuse.Value)
                        {
                            var err = new RFLogError();
                            err.Type = ErrorType.FuseReuse;
                            err.Level = EnumErrorLevel.Warning;
                            err.Message = $"FuseName: {fuseitem.FuseName} already used in previous address : {fuseitem.Address}-{fuseitem.LSB}-{fuseitem.MSB}";
                            RFLogManager.Push(err, LogName, fuseitem.RowNum, "EFUSENAME,ADDRESS,LSB,MSB");
                        }
                    }
                }
                CheckTrim(patSubName, trimInfoList);
                #endregion
            }
            return result;
        }

        private void ProcessBenchLogOperation(BenchLogItem loginfo, string logDirect, Dictionary<string, Register> tmpregisters, MappingItem patternName, Dictionary<string, Register> registers, List<BenchLogItem> trimLogs, List<BenchLogItem> usedTrimLogs, ref List<string> trimInfoList, ref string patSubName)
        {
            switch (loginfo.Operation.ToLower())
            {
                case "patname":
                    trimInfoList = new List<string> { "trim_start", "trim_end", "write_trim" };
                    patSubName = loginfo.MSB.ToString();
                    break;
                case "calc":
                    CheckCalc(loginfo.PostProcess, loginfo.TestName, loginfo.RowNum);
                    CheckAllStoreNamesDup(loginfo.PostProcess, loginfo.RowNum);
                    break;
                case "reg_read":
                case "read":
                case "lut":
                case "wait":
                case "forcev":
                case "forcei":
                    break;
                case "sequence order":
                case "sequence_order":
                    RFLogError err = UpdateTableValue(logDirect, loginfo.Interface, tmpregisters, patternName.Pattern);
                    if (err != null)
                    { RFLogManager.Push(err, LogName, loginfo.RowNum, "INTERFACE"); }
                    break;
                case "reg_write":
                    CheckWrite(loginfo, registers, true);
                    break;
                case "write_trim":
                    trimLogs.Add(loginfo);
                    usedTrimLogs.Add(loginfo);
                    trimInfoList.Remove(loginfo.Operation.ToLower());
                    CheckWrite(loginfo, registers);
                    CheckFieldVal(loginfo.FieldVal, loginfo.RowNum);
                    break;
                case "trim_start":
                case "trim_end":
                    CheckLimit(loginfo.LoLimit, loginfo.HighLimit, loginfo.RowNum);
                    trimInfoList.Remove(loginfo.Operation.ToLower());
                    break;
                case "write":
                    CheckWrite(loginfo, registers);
                    break;
                case "write_src":
                    HandleWriteSrcOperation(loginfo, registers);
                    break;
                case "wait_for":
                    CheckLimit(loginfo.LoLimit, loginfo.HighLimit, loginfo.RowNum);
                    CheckFieldVal(loginfo.FieldVal, loginfo.RowNum);
                    break;
                case "read_cap":
                    HandleReadCapOperation(loginfo);
                    break;
                case "read_compare":
                    CheckFieldVal(loginfo.FieldVal, loginfo.RowNum);
                    break;
                case "otp_write":
                    HandleOtpWriteOperation(loginfo);
                    break;
                case "spac_mem_read":
                    CheckAllStoreNamesDup(loginfo.PostProcess, loginfo.RowNum);
                    break;
                default:
                    HandleDefaultOperation(loginfo);
                    break;
            }
        }

        private void HandleWriteSrcOperation(BenchLogItem loginfo, Dictionary<string, Register> registers)
        {
            if ((Regex.IsMatch(loginfo.Default, "^0x", RegexOptions.IgnoreCase) || int.TryParse(loginfo.Default, out int _) || string.IsNullOrEmpty(loginfo.Default)) &&
                !string.IsNullOrEmpty(loginfo.FuseName))
            {
                loginfo.Default = loginfo.FuseName;
            }
            CheckWrite(loginfo, registers);
        }

        private void HandleReadCapOperation(BenchLogItem loginfo)
        {
            CheckPatBits(loginfo);
            CheckLimit(loginfo.LoLimit, loginfo.HighLimit, loginfo.RowNum);
            CheckTestNameDup(loginfo.TestName, loginfo.RowNum);
            string storename = CheckAllStoreNamesDup(loginfo.PostProcess, loginfo.RowNum);

            string srcReg = GetDigSrcRegisterName(loginfo);
            foreach (KeyValuePair<string, List<string>> srcWarnRec in _srcwarnrecords)
            {
                if (srcWarnRec.Key == srcReg)
                {
                    srcWarnRec.Value.Add(storename);
                }
            }
        }

        private void HandleOtpWriteOperation(BenchLogItem loginfo)
        {
            CheckFuseNameExist(loginfo.RowNum, loginfo.Interface, "INTERFACE");
            foreach (string reckey in _srcwarnrecords.Keys)
            {
                if (_srcwarnrecords.ContainsKey(reckey) && _srcwarnrecords[reckey].Contains(loginfo.Address))
                {
                    _srcwarnrecords[reckey].Add(loginfo.Interface);
                }
            }
        }

        private void HandleDefaultOperation(BenchLogItem loginfo)
        {
            CheckLimit(loginfo.LoLimit, loginfo.HighLimit, loginfo.RowNum);
            if (Regex.IsMatch(loginfo.Operation, "meas|INJECT|autogen", RegexOptions.IgnoreCase))
            {
                //Check measure pin exist
                if (loginfo.Operation.ToLower().Contains("meas") || loginfo.Operation.ToLower().Contains("inject"))
                {
                    //if pinmap contains pin TBD
                    CheckPinExists(loginfo.Interface, loginfo.RowNum);
                }
                else
                {
                    //case 2: AUTOGEN, //TE_OPERATION WiMeas Pin=ARF_PADIO_TXOUT_5G_0//TestType = PSAT;IQMM;LOFT//freq=5183MHz&5177MHz&5180MHz//rlevl=0dBm // MeasName = ARF_14_WL5GRXC0_PSAT5180_ARFPADIOTXOUT5G0_MEASP_X_X_X_X_NV_VXXX,,,,,,,,,,,,,,

                    CheckTestNameDup(loginfo.TestName, loginfo.RowNum);
                    string regPin = @"(?<MeasType>\w+)\sPin\s*=\s*(?<pin>[\w\:\,]+)";
                    string regTestType = @"TestType\s*=\s*(?<TestType>[\w\;]+)";
                    string pinName = Regex.Match(loginfo.Interface, regPin, RegexOptions.IgnoreCase).Groups["pin"].Value;

                    //if pinmap contains pin TBD
                    CheckPinExists(pinName, loginfo.RowNum);

                    //check testtype if use wimeas
                    string measType = Regex.Match(loginfo.Interface, regPin, RegexOptions.IgnoreCase).Groups["MeasType"].Value;
                    string testType = Regex.Match(loginfo.Interface, regTestType, RegexOptions.IgnoreCase).Groups["TestType"].Value;
                    CheckMeasType(measType, testType, loginfo.RowNum);

                    CheckAllStoreNamesDup(loginfo.PostProcess, loginfo.RowNum);
                }
            }
        }

        private void CheckTrim(string patname, List<string> trimList)
        {
            if (trimList.Count > 0 && trimList.Count < 3)
            {
                var err = new RFLogError();
                err.Type = ErrorType.TrimInfoNotComplete;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("PatName:{0}, need to fill trim syntax to avoid checking error: {1} \r\n", patname, string.Join(",", trimList));
                RFLogManager.Push(err, LogName, 0, "OPERATION");
            }
        }

        private void CheckPinExists(string pinName, int row)
        {
            if (Pinmap == null)
            {
                return;
            }

            IEnumerable<string> pins = Regex.Split(pinName, @"[:,]", RegexOptions.IgnoreCase).Where(p => !string.IsNullOrEmpty(p));
            foreach (string pin in pins)
            {
                if (!Pinmap.IsPinExist(pin) && !Pinmap.IsGroupExist(pin))
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.PinMissing;
                    if (Pinmap.PinList.Any(p => p.PinName.StartsWith(pin + "_", StringComparison.OrdinalIgnoreCase)))
                    {
                        IEnumerable<string> pmpins = Pinmap.PinList.Where(p => p.PinName.StartsWith(pin + "_", StringComparison.OrdinalIgnoreCase)).Select(pmpin => pmpin.PinName);
                        err.Level = EnumErrorLevel.Warning;
                        err.Message = string.Format("Please check pin {0} is Exist, PinMap : {1}", pin, string.Join(",", pmpins));
                    }
                    else
                    {
                        err.Level = EnumErrorLevel.Error;
                        err.Message = string.Format("Please check pin {0} is Exist", pin);
                    }
                    RFLogManager.Push(err, LogName, row, "INTERFACE");
                }
            }
        }

        private RFLogError UpdateTableValue(string path, string refinterface, Dictionary<string, Register> totalRegisters, string payload)
        {
            RFLogError err = null;
            string refLog = refinterface.Split('#')[0];
            if (string.IsNullOrEmpty(refLog))
            {
                err = new RFLogError();
                err.Type = ErrorType.EmptyInit;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("Empty Sequence order");
                return err;
            }
            bool isFindTarget = true;
            foreach (string logFile in Directory.GetFiles(path, "*.csv"))
            {

                if (!Path.GetFileNameWithoutExtension(logFile).Equals(refinterface))
                {
                    continue;
                }

                bool isFindFile = true;
                List<BenchLogFile> loginfos = ReadLogFile(logFile, false);
                foreach (BenchLogFile loginfo in loginfos)
                {
                    List<BenchLogItem> logs = loginfo.Logs;
                    foreach (BenchLogItem log in logs)
                    {
                        switch (log.Operation.ToLower())
                        {
                            case "patname":
                                if (!refLog.Equals(log.MSB))
                                {
                                    isFindTarget = false;
                                }

                                break;
                            case "sequence order":
                            case "sequence_order":
                                RFLogError error = UpdateTableValue(path, log.Interface, totalRegisters, payload);
                                if (error != null)
                                { RFLogManager.Push(error, LogName, log.RowNum, "INTERFACE"); }
                                break;
                            case "reg_write":
                                UpdateRegisterValue(log, totalRegisters, log.Default, true);
                                break;
                            case "write":
                                UpdateRegisterValue(log, totalRegisters, log.Default, false);
                                break;
                            case "write_src":
                                UpdateRegisterValue(log, totalRegisters, log.Default, false);
                                break;

                            default:
                                break;
                        }

                    }
                    if (isFindTarget)
                    {
                        break;
                    }
                }
                if (!isFindTarget || !isFindFile)
                {
                    err = new RFLogError();
                    err.Type = ErrorType.InitNotFound;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("Sequence Order Not Found : {0}", refLog);
                }
            }
            return err;
        }


        private Register UpdateRegisterValue(BenchLogItem log, Dictionary<string, Register> registers, string relatedVar, bool isReg)
        {
            if (isReg)
            {
                ;
            }

            string address = log.Address;

            if (!string.IsNullOrEmpty(relatedVar) && (Regex.IsMatch(relatedVar, "^0x", RegexOptions.IgnoreCase) || double.TryParse(relatedVar, out _)))
            {
                relatedVar = "";
            }

            if (string.IsNullOrEmpty(relatedVar))
            {
                _ = log.FuseName;
            }

            int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
            bool isNeedShift = PatBit == 64 && addrQuot % 2 != 0;
            int shiftbit = 0;
            if (isNeedShift)
            {
                shiftbit = 32;
            }

            Register result = registers.ContainsKey(address) ? registers[address] : null;

            if (result == null)
            {
                result = CreateRegister(log, log.Default, isNeedShift);
                registers.Add(result.Address, result);
            }

            if (string.IsNullOrEmpty(log.RegData))
            {
                #region marked

                //var binStr = _ConvertValToBin(data, int.Parse(result.Length), false);
                //if (isReg)
                //{
                //    var setValue_fail = false;
                //    var binArr = binStr.ToCharArray();
                //    for (int i = 0; i < 32; i++)
                //    {
                //        if (result.Data.Substring(i, 1) == "D")
                //        {
                //            binArr[i] = 'D'; 
                //        }
                //    }
                //    result.Data = new string(binArr);


                //    foreach (var field in result.Fields)
                //    {
                //        var tmp = binStr.Substring(binStr.Length - field.BitSize, field.BitSize);
                //        binStr = binStr.Substring(0, binStr.Length - field.BitSize);
                //        setValue_fail = setValue_fail || field.SetValue(tmp);
                //    }
                //    if (setValue_fail)
                //    {
                //        var err = new LogError();
                //        err.Type = ErrorType.DataNotUpdateByValue;
                //        err.Level = ErrorLevel.Error;
                //        err.Message = string.Format("Value : {0} could not update \r\nbecause already set by Write_src \r\nAddress:{1}/MSB:{2}/LSB:{3}"
                //            , log.Data, log.Address, log.MSB, log.LSB);
                //        LogManager.Push(err, LogName, log.RowNum, "Address,LSB,MSB");
                //    }
                //}
                //else
                //{
                //var binStr = _ConvertValToBin(data, int.Parse(result.Length), false);
                //var regField = @"(?<name>\w+)(?<partial>\[.*\])*";
                //    var tarField = Regex.Match(register, regField, RegexOptions.IgnoreCase).Groups["name"].Value;
                //    var partialInfo = Regex.Match(register, regField, RegexOptions.IgnoreCase).Groups["partial"].Value;
                //    var regdataArr = result.Data.ToCharArray();
                //    var binArr = binStr.ToCharArray();
                //    var index = 0;
                //    for (var i = 31 - log.LSB; i >= 31 - log.MSB; i--)
                //    {
                //        if (!string.IsNullOrEmpty(log.FuseName))
                //            regdataArr[i] = 'D';
                //        else
                //        {
                //            regdataArr[i] = binArr[binArr.Length - 1 - index];
                //        }
                //        index++;
                //    }
                //    result.Data = new string(regdataArr);
                //}

                #endregion
            }
            else
            {
                ApplyRegisterData(log, result, shiftbit, isNeedShift);
            }

            return result;
        }

        private void ApplyRegisterData(BenchLogItem log, Register result, int shiftbit, bool isNeedShift)
        {
            char[] sourceStr = result.Data.ToCharArray();
            char[] regDataStr = ConvertValToBin(log.RegData, int.Parse(result.Length), false, isNeedShift).ToCharArray();
            bool setValue_fail = false;

            for (int bit = log.LSB + shiftbit; bit <= log.MSB + shiftbit; bit++)
            {
                int idx = PatBit - 1 - bit;

                if (sourceStr[idx] == 'D')
                {
                    setValue_fail = true;
                    continue;
                }

                sourceStr[idx] = regDataStr[idx];
            }

            if (setValue_fail && log.Operation.Equals("Write", StringComparison.OrdinalIgnoreCase))
            {
                var err = new RFLogError();
                err.Type = ErrorType.DataNotUpdateByValue;
                err.Level = _isOverWrite ? EnumErrorLevel.Warning : EnumErrorLevel.Error;
                if (_isOverWrite)
                {
                    err.Type = ErrorType.DataUpdateByValue;
                    err.Message = string.Format("Value : Address: {0}/MSB:{1}/LSB:{2} got updated with fixed value instead of WRITE_SRC value."
                            , log.Address, log.MSB, log.LSB);
                }
                else
                {
                    err.Type = ErrorType.DataNotUpdateByValue;
                    err.Message = string.Format("Value : {0} could not update \r\nbecause already set by Write_src \r\nAddress:{1}/MSB:{2}/LSB:{3}"
                            , log.FieldVal, log.Address, log.MSB, log.LSB);
                }
                RFLogManager.Push(err, LogName, log.RowNum, "ADDRESS,LSB,MSB");
            }

            if (!string.IsNullOrEmpty(log.FuseName))
            {
                for (int i = log.LSB + shiftbit; i <= log.MSB + shiftbit; i++)
                {
                    sourceStr[PatBit - 1 - i] = 'D';
                }
            }
            result.Data = new string(sourceStr);
        }

        private void CompareWithEfuse(BenchLogItem log)
        {
            if (string.IsNullOrEmpty(log.FuseName))
            {
                return;
            }

            List<string> fuse = CheckFuseNameExist(log.RowNum, log.FuseName, "efuseName");

            if (fuse != null)
            {
                int size = log.MSB - log.LSB + 1;
                if (int.Parse(fuse[3]) != size)
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.FuseSizeMisMatch;
                    err.Level = _isFW ? EnumErrorLevel.Warning : EnumErrorLevel.Error;
                    err.Message = string.Format("FuseName : {0} \r\nBenchLog : {1}, FuseTable : {2}", log.FuseName, size, fuse[3]);
                    RFLogManager.Push(err, LogName, log.RowNum, "MSB,LSB,EFUSENAME");
                }
            }
        }

        private List<string> CheckFuseNameExist(int rownum, string fusename, string header)
        {
            List<string> fuse =
                EfusebitDefRows.FirstOrDefault(p => p[0].Equals(fusename, StringComparison.OrdinalIgnoreCase));

            bool storename =
                LogStoreName.Contains(fusename) || LogCalcStoreName.Contains(fusename) || LogMeasStoreName.Contains(fusename);

            if (fuse == null && !storename)
            {
                var err = new RFLogError();
                err.Type = ErrorType.FuseCategoryNotFound;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("FuseName : {0} \r\nNot found in efuse bitdefinition table", fusename);
                RFLogManager.Push(err, LogName, rownum, header);
            }

            return fuse;
        }

        private static string ConvertValToBin(string data, int size, bool isNeedReverse, bool isShiftToHight = false)
        {
            string result = Regex.Match(data, @"(0x)*(?<data>\w+)", RegexOptions.IgnoreCase).Groups["data"].Value;
            //Convert by Hexdecimal
            try
            {
                if (string.IsNullOrEmpty(data))
                {
                    ;
                }
                else if (data.Substring(data.Length - 1, 1) != "0")
                {
                    ;
                }
            }
            catch (Exception)
            {
                ;
            }

            if (Regex.IsMatch(data, @"0x(?<data>\w+)", RegexOptions.IgnoreCase))
            {
                result = string.Join(string.Empty, result.Select(c =>
                Convert.ToString(Convert.ToInt64(c.ToString(), 16), 2).PadLeft(4, '0')));
            }
            else
            {
                result = string.Join(string.Empty, result.Select(c =>
                Convert.ToString(Convert.ToInt64(c.ToString(), 10), 2)));
            }

            if (isShiftToHight && result.Length <= 32)
            {
                result = result.PadLeft(32, '0').PadRight(64, '0');
            }
            else
            {
                result = result.Length < size ?
                    result = result.PadLeft(size, '0') :
                    result = result.Substring(result.Length - size, size);
            }

            if (isNeedReverse)
            {
                char[] resultArr = result.ToCharArray();
                Array.Reverse(resultArr);
                return new string(resultArr);
            }
            return result;
        }

        private void CheckCalc(string postprocess, string testname, int row)
        {
            List<string> args = new List<string>();
            string calcstorename = "";
            foreach (string subline in postprocess.Trim(',').Replace("=", ":").Split(';'))
            {
                string subkey = subline.Split(':')[0];
                string subvalue = subline.Split(':')[1];
                if (subkey.Trim().ToLower() == "calc")
                {
                    string allargs = subvalue.Substring(subvalue.IndexOf("(") + 1).TrimEnd(')');
                    args.AddRange(allargs.Split(','));
                }
                else if (subkey.Trim().ToLower() == "calcstorename")
                {
                    calcstorename = subvalue.Trim();
                }
                else if (subkey.Trim().ToLower() == "storename")
                {
                    // 260401_Avoid case StoreName:5gauxadccaltrimhminch1;Calc:auxADC_infer_caltrimh_minrange(DOUT_0V7_WL5GAUX1_H,DOUT_0V85_WL5GAUX1_H);CalcStoreName:5gauxadc_ch1_caltrimh_minrange
                    var err = new RFLogError();
                    err.Type = ErrorType.StroeNameErrorWtihCalc;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = $"Calc can not use \"StoreName\" {subline}";
                    RFLogManager.Push(err, LogName, row, "POSTPROCESS");
                }
            }

            foreach (string arg in args)
            {
                foreach (KeyValuePair<string, string> srcRec in _srcrecords)
                {
                    if (srcRec.Value == arg)
                    {
                        _srcwarnrecords[srcRec.Key].AddRange(calcstorename.Split(','));
                    }
                }

                if (!int.TryParse(arg, out _) &&
                    !decimal.TryParse(arg, out _) &&
                    !double.TryParse(arg, out _) &&
                    Regex.IsMatch(arg, @"^[a-zA-Z0-9_]+$") &&
                    !LogStoreName.Contains(arg) && !LogCalcStoreName.Contains(arg) && EfusebitDefRows.FirstOrDefault(p => p[0].Equals(arg, StringComparison.OrdinalIgnoreCase)) == null)
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.CalcArgUninitialize;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = $"Arg : {arg} uses uninitialized names";
                    RFLogManager.Push(err, LogName, row, "POSTPROCESS");
                }
            }

            CheckTestNameDup(testname, row);

            string[] csNames = calcstorename.Split(',');
            string[] tNames = testname.Split(',');

            if (tNames.Length > 1 && csNames.Length != tNames.Length)
            {
                var err = new RFLogError();
                err.Type = ErrorType.CalcStoreNameTestNameErrorNum;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("CalcStoreName : {0} != TestName : {1}", csNames.Length, tNames.Length);
                RFLogManager.Push(err, LogName, row, "CALCSTORENAME,TESTNAME");
            }
            /// Draco Perform a pre-check to verify whether the Python library BenchLog is available.
            if (FunctionMaps.Count != 0)
            {
                //    foreach(var csName in csNames)
                //    {
                //        if(!FunctionMaps.ContainsKey(csName))
                //        {
                //            var err = new RFLogError();
                //            err.Type = ErrorType.CalcFunMismatchWithPyhtonFile;
                //            err.Level = ErrorLevel.Error;
                //            err.Message = string.Format("CalcFunction: {0} cannot be found in the Python file", csName);
                //            RFLogManager.Push(err, LogName, row, "CALCSTORENAME");
                //        }
                //    }
            }

        }

        private string CheckAllStoreNamesDup(string postprocess, int row)
        {
            string storeName = "";
            string calcStoreName = "";
            string measStoreName = "";
            List<string> fullStoreNames = new List<string>();
            foreach (string subline in postprocess.Trim(',').Replace("=", ":").Split(';'))
            {
                if (subline.Split(':')[0].Trim().ToLower() == "storename")
                {
                    storeName = subline.Split(':')[1].Trim();
                }

                if (subline.Split(':')[0].Trim().ToLower() == "calcstorename")
                {
                    calcStoreName = subline.Split(':')[1].Trim();
                }

                if (subline.Split(':')[0].Trim().ToLower() == "measstorename")
                {
                    measStoreName = subline.Split(':')[1].Trim();
                }
            }

            if (!string.IsNullOrEmpty(storeName))
            {
                CheckStoreNameDup(storeName, row);
                fullStoreNames.Add(storeName);
            }
            if (!string.IsNullOrEmpty(calcStoreName))
            {
                foreach (string csn in calcStoreName.Split(','))
                {
                    CheckCalcStoreNameDup(csn, row);
                    fullStoreNames.Add(csn);
                }
            }
            if (!string.IsNullOrEmpty(measStoreName))
            {
                CheckMeasStoreNameDup(measStoreName, row);
                fullStoreNames.Add(measStoreName);
            }

            foreach (string fsn in fullStoreNames)
            {
                CheckFullStoreNameDup(fsn, row);
            }

            return storeName;
        }

        private void CheckTestNameDup(string testname, int row)
        {
            foreach (string tn in testname.Split(','))
            {
                if (!string.IsNullOrEmpty(tn))
                {
                    if (LogTestName.Contains(tn))
                    {
                        var err = new RFLogError();
                        err.Type = ErrorType.DuplicateTestNameInside;
                        err.Level = EnumErrorLevel.Error;
                        err.Message = string.Format("TestName : {0} , Already used previously", tn);
                        RFLogManager.Push(err, LogName, row, "TESTNAME");
                    }
                    else
                    {
                        LogTestName.Add(tn);
                    }
                }
            }
        }

        private void CheckStoreNameDup(string storename, int row)
        {
            if (LogStoreName.Contains(storename))
            {
                var err = new RFLogError();
                err.Type = ErrorType.DuplicateStoreNameInside;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("StoreName : {0} , Already used previously", storename);
                RFLogManager.Push(err, LogName, row, "POSTPROCESS");
            }
            else
            {
                LogStoreName.Add(storename);
            }
        }

        private void CheckCalcStoreNameDup(string calcstorename, int row)
        {
            if (LogCalcStoreName.Contains(calcstorename))
            {
                var err = new RFLogError();
                err.Type = ErrorType.DuplicateCalcStoreNameInside;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("CalcStoreName : {0} , Already used previously", calcstorename);
                RFLogManager.Push(err, LogName, row, "POSTPROCESS");
            }
            else
            {
                LogCalcStoreName.Add(calcstorename);
            }
        }

        private void CheckMeasStoreNameDup(string measstorename, int row)
        {
            if (LogMeasStoreName.Contains(measstorename))
            {
                var error = new RFLogError();
                error.Type = ErrorType.DuplicateMeasStoreNameInside;
                error.Level = EnumErrorLevel.Error;
                error.Message = string.Format("MeasStoreName : {0} , Already used previously", measstorename);
                RFLogManager.Push(error, LogName, row, "POSTPROCESS");
            }
            else
            {
                LogMeasStoreName.Add(measstorename);
            }
        }

        private void CheckFullStoreNameDup(string fullstorename, int row)
        {
            if (LogFullStoreName.Contains(fullstorename))
            {
                var error = new RFLogError();
                error.Type = ErrorType.DuplicateFullStoreNameInside;
                error.Level = EnumErrorLevel.Error;
                error.Message = string.Format("Store Name : {0} , Already used in all Store Name", fullstorename);
                RFLogManager.Push(error, LogName, row, "POSTPROCESS");
            }
            else
            {
                LogFullStoreName.Add(fullstorename);
            }
        }

        private void CheckMeasType(string meastype, string testtype, int row)
        {
            if (meastype.Equals("WiMeas", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(testtype))
                {
                    var error = new RFLogError();
                    error.Type = ErrorType.TestTypeMissing;
                    error.Level = EnumErrorLevel.Error;
                    error.Message = string.Format("Not Found test type by using WiMeas");
                    RFLogManager.Push(error, LogName, row, "INTERFACE");
                }
            }
        }

        private void CheckPatBits(BenchLogItem log)
        {
            if (log.RegData.Length > 10 && PatBit == 32)
            {
                var err = new RFLogError();
                err.Type = ErrorType.PatternBitError;
                err.Level = EnumErrorLevel.Error;
                err.Message = $"Error occurs using 64 bits REG_DATA : {log.RegData} in 32bits pattern";
                RFLogManager.Push(err, LogName, log.RowNum, "EFUSENAME,ADDRESS,LSB,MSB");
            }
        }

        private void CheckWrite(BenchLogItem log, Dictionary<string, Register> registers, bool isReg = false)
        {

            CheckPatBits(log);

            CompareWithEfuse(log);

            UpdateRegisterValue(log, registers, log.Default, isReg);

            StoreEfuseDuplicateUsage(log);


            if (log.Operation.Equals("WRITE_SRC", StringComparison.OrdinalIgnoreCase))
            {
                string srcReg = GetDigSrcRegisterName(log);

                if (!_srcrecords.ContainsKey(srcReg))
                {
                    _srcrecords.Add(srcReg, log.Default);
                    _srcwarnrecords.Add(srcReg, new List<string>());
                }
                else
                {
                    if (_srcrecords[srcReg] != log.Default)
                    {
                        var err = new RFLogError();
                        err.Type = ErrorType.FieldUseDiffFuse;
                        err.Level = EnumErrorLevel.Warning;
                        if (_srcwarnrecords[srcReg].Contains(log.Default))
                        {
                            err.Message = string.Format("Address : {0}, LSB : {1}, MSB : {2} Fuse name be changed by Read_Cap, Calc or OTP_Write. \r\nPreviously: {3}, Current: {4}"
                                , log.Address, log.LSB, log.MSB, _srcrecords[srcReg], log.Default);
                        }
                        else
                        {
                            err.Message = string.Format("Address : {0}, LSB : {1}, MSB : {2} Use different fuse value. \r\nPreviously: {3}, Current: {4}"
                                , log.Address, log.LSB, log.MSB, _srcrecords[srcReg], log.Default);
                        }
                        RFLogManager.Push(err, LogName, log.RowNum, "EFUSENAME,ADDRESS,LSB,MSB");
                    }
                }
            }
        }
        private static string GetDigSrcRegisterName(BenchLogItem log)
        {
            string srcreg = $"Addr{log.Address}_M{log.MSB}_L{log.LSB}";
            string result = srcreg;
            return result;
        }

        private void StoreEfuseDuplicateUsage(BenchLogItem log)
        {
            if (string.IsNullOrEmpty(log.FuseName))
            {
                return;
            }

            if (!_rEuseDictionary.ContainsKey(log.FuseName))
            {
                _rEuseDictionary.Add(log.FuseName, new List<BenchLogItem>());
            }

            _rEuseDictionary[log.FuseName].Add(log);
        }

        private void CheckFieldVal(string fieldval, int row)
        {
            if ((fieldval.StartsWith("0x") && ulong.TryParse(fieldval.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out _)) ||
                ulong.TryParse(fieldval, out _))
            {
                return;
            }
            else
            {
                var err = new RFLogError();
                err.Type = ErrorType.FieldValNotValue;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("Field_Val is not valid : {0}", fieldval);
                RFLogManager.Push(err, LogName, row, "FIELD_VAL");
            }
        }

        private void CheckLimit(string lolim, string hilim, int row)
        {
            if (!string.IsNullOrEmpty(lolim) && !double.TryParse(lolim, out _))
            {
                if (!lolim.Equals("NA") && !lolim.StartsWith("0x"))
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.LimitNotValue;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("Limit is not valid : {0}", lolim);
                    RFLogManager.Push(err, LogName, row, "LOLIMIT");
                }
            }

            if (!string.IsNullOrEmpty(hilim) && !double.TryParse(hilim, out _))
            {
                if (!hilim.Equals("NA") && !hilim.StartsWith("0x"))
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.LimitNotValue;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("Limit is not valid : {0}", hilim);
                    RFLogManager.Push(err, LogName, row, "HILIMIT");
                }
            }
        }
        public static bool IsFileLock(ref string path)
        {
            try
            {
                var sr = new StreamReader(path);

                sr.Close();

            }
            catch (IOException)
            {
                string ext = Path.GetExtension(path);
                string file = Path.GetFileNameWithoutExtension(path);
                file = file + "_dummy";
                path = path.Replace(Path.GetFileName(path), file + ext);
                return true;
            }
            return false;
        }

        private List<BenchLogFile> ReadLogFile(string originpath, bool isNeedCheck = true, bool isFullSweep = false)
        {
            var result = new List<BenchLogFile>();
            var loginfo = new BenchLogFile();
            string logName = Path.GetFileNameWithoutExtension(originpath);
            string reg_meas = @"//TE_OPERATION\s+//(?<InstType>\w+)\s+(?:Pin|DisconnectPins)\s*=\s*(?<PinName>.*?)(?=\s*//|$)";

            List<ChannelMapSheet> channelMapSheet = new List<ChannelMapSheet>();
            var genInstrumentSetupTable = new GenInstrumentSetupTable();
            genInstrumentSetupTable.Init(channelMapSheet);
            var path = genInstrumentSetupTable.Writer.InstTypeSetting.Where(x => x.Path.Equals("to_SRC", StringComparison.OrdinalIgnoreCase)).ToList();
            if (!string.IsNullOrEmpty(LocalSpecs.PinMap))
            {
                Pinmap = new ReadPinMapSheet().GetSheet(LocalSpecs.PinMap);
            }
            loginfo.Offset = -1;
            bool isNeedDelete = false;
            string tmp = originpath;
            while (IsFileLock(ref tmp))
            {
                if (!File.Exists(tmp))
                {
                    File.Copy(originpath, tmp);
                    isNeedDelete = true;
                }
                ;
            }

            string line = "";
            var sr = new StreamReader(tmp);
            bool isFirst = true;
            int row = 0;
            while ((line = sr.ReadLine()) != null)
            {
                loginfo.RawDatas.Add(line);
                row++;
                try
                {
                    ProcessLogLine(line, ref loginfo, result, row, logName, isNeedCheck, isFullSweep, ref isFirst, reg_meas, path);
                }
                catch (Exception)
                {
                    ;
                }
            }
            sr.Close();
            if (isNeedDelete)
            {
                File.Delete(tmp);
            }

            return result;
        }

        private void ProcessLogLine(string line, ref BenchLogFile loginfo, List<BenchLogFile> result, int row, string logName, bool isNeedCheck, bool isFullSweep, ref bool isFirst, string regMeas, List<RfLib.InstrumentSetup.InstrumentTypeData.InstrumentTypeSetting> path)
        {
            if (string.IsNullOrEmpty(line))
            {
                return;
            }
            CheckDuplicatePatNameRow(loginfo, row, logName);
            CheckLogHeader(line, loginfo, row, logName, isNeedCheck);

            var log = new BenchLogItem(line, loginfo, row, isFullSweep, ref isFirst);

            log.RowNum = row;

            Match measInfo = Regex.Match(log.Interface, regMeas, RegexOptions.IgnoreCase);

            if (log.Operation.Equals("PRINT", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (string.IsNullOrEmpty(log.Operation))
            {
                return;
            }

            if (Regex.IsMatch(log.Operation, "PATNAME", RegexOptions.IgnoreCase))
            {
                HandlePatNameLine(log, ref loginfo, result, row, logName, isNeedCheck);

                return;
            }
            if (Regex.IsMatch(log.Operation, "OPERATION|print", RegexOptions.IgnoreCase))
            {
                return;
            }

            AppendLogByOperation(log, loginfo, line, row, logName, isFullSweep, ref isFirst, measInfo, path);
        }

        private static void CheckDuplicatePatNameRow(BenchLogFile loginfo, int row, string logName)
        {
            if (loginfo.RawDatas.Count != 2 &&
                loginfo.RawDatas[loginfo.RawDatas.Count - 1].IndexOf("patname", StringComparison.OrdinalIgnoreCase) >= 0 &&
                loginfo.RawDatas[loginfo.RawDatas.Count - 2].IndexOf("patname", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var err = new RFLogError();
                err.Type = ErrorType.NoPatContent;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("NoPatContent : {0}", loginfo.RawDatas[loginfo.RawDatas.Count - 1]);
                RFLogManager.Push(err, logName, row - 1, "");
            }
        }

        private static void CheckLogHeader(string line, BenchLogFile loginfo, int row, string logName, bool isNeedCheck)
        {
            if (loginfo.Offset == -1 && line.ToUpper().Contains("OPERATION"))
            {
                loginfo.Offset =
                    line.Split(',')
                        .ToList()
                        .FindIndex(p => p.Equals("OPERATION", StringComparison.OrdinalIgnoreCase));
                //BenchLogItem.ResetIndex();
                List<string> headerissues = loginfo.GetHeaderAndCheck(line);
                if (headerissues.Count > 0 && isNeedCheck)
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.HeaderMissing;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("Missing : {0}", string.Join(",", headerissues));
                    RFLogManager.Push(err, logName, row, "");
                }
            }
        }

        private void HandlePatNameLine(BenchLogItem log, ref BenchLogFile loginfo, List<BenchLogFile> result, int row, string logName, bool isNeedCheck)
        {
            List<string> logsplit = log.RawDatas;
            string tmpName = logsplit[1] + "_" + logsplit[2] + "_" + logsplit[3] + "_" + logsplit[4] + "_" + logsplit[5];

            if (PatNames.Contains(tmpName) && isNeedCheck)
            {
                var err = new RFLogError();
                err.Type = ErrorType.DuplicatePatternNameInside;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("PatName: {0}, Already used previously.", tmpName);
                RFLogManager.Push(err, logName, row, "OPERATION");
            }
            PatNames.Add(tmpName);

            loginfo = loginfo.Copy();
            loginfo.LogFile = logName;
            var patNameIssue = new List<string> { };
            loginfo.PatternSubName = logsplit[3];
            CheckPatternName(row, logName, loginfo.PatternSubName, logsplit[4]);
            loginfo.Type = logsplit[2];
            loginfo.Interface = logsplit[1];
            loginfo.Version = string.IsNullOrEmpty(logsplit[4]) ? "1" : logsplit[4];
            loginfo.DateCode = string.IsNullOrEmpty(logsplit[5]) ? "" : logsplit[5];

            if (string.IsNullOrEmpty(loginfo.PatternSubName))
            {
                patNameIssue.Add("PatternName");
            }

            if (string.IsNullOrEmpty(loginfo.Type))
            {
                patNameIssue.Add("PatternType");
            }
            if (patNameIssue.Count > 0 && isNeedCheck)
            {
                var err = new RFLogError();
                err.Type = ErrorType.PatNameIssue;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("Missing : {0}", string.Join(",", patNameIssue));
                RFLogManager.Push(err, logName, log.RowNum, "");
            }
            result.Add(loginfo);
        }

        private void AppendLogByOperation(BenchLogItem log, BenchLogFile loginfo, string line, int row, string logName, bool isFullSweep, ref bool isFirst, Match measInfo, List<RfLib.InstrumentSetup.InstrumentTypeData.InstrumentTypeSetting> path)
        {
            if (log.Operation.Equals("WAIT", StringComparison.OrdinalIgnoreCase))
            {
                loginfo.Logs.Add(log);
            }
            else if (log.Operation.Equals("Read_cap", StringComparison.OrdinalIgnoreCase))
            {
                if (loginfo.Logs.Count > 0)
                {
                    BenchLogItem logCap = loginfo.Logs.Last(p => !p.Operation.Equals("calc", StringComparison.OrdinalIgnoreCase));

                    loginfo.Logs.Add(log);
                }
                else
                {
                    loginfo.Logs.Add(log);
                }
            }

            else if (Regex.IsMatch(log.Operation, "write_src", RegexOptions.IgnoreCase))
            {

                if (loginfo.Logs.Count > 0)
                {
                    loginfo.Logs.Add(log);
                }
                else
                {
                    loginfo.Logs.Add(log);
                }
            }
            else if (Regex.IsMatch(log.Operation, "wait_for", RegexOptions.IgnoreCase))
            {
                loginfo.Logs.Add(log);
                string newLine = Regex.Replace(line, "wait_for", "read_cap", RegexOptions.IgnoreCase);
                log = new BenchLogItem(newLine, loginfo, row, isFullSweep, ref isFirst);
                loginfo.Logs.Add(log);
            }
            else if (measInfo.Success)
            {
                loginfo.Logs.Add(log);
                CheckSourceFrequencyRange(log, measInfo, row, logName, path);
            }
            else
            {
                loginfo.Logs.Add(log);
            }
        }

        private void CheckSourceFrequencyRange(BenchLogItem log, Match measInfo, int row, string logName, List<RfLib.InstrumentSetup.InstrumentTypeData.InstrumentTypeSetting> path)
        {
            string pinname = measInfo.Groups["PinName"].Value;
            string pintype = measInfo.Groups["InstType"].Value;

            Match freqMatch = Regex.Match(log.Interface, @"//freq=(?<Freq>[\d.]+)", RegexOptions.IgnoreCase);

            double? frequency = null;

            if (freqMatch.Success)
            {
                frequency = double.Parse(freqMatch.Groups["Freq"].Value);
            }

            if (pinname.Contains("::"))
            {
                string[] allpins = pinname.Split(new string[] { "::" }, StringSplitOptions.None);
                foreach (string pin in allpins)
                {
                    if (Pinmap.PinList.Any(p => string.Equals(p.PinName, pin + "_SRC", StringComparison.OrdinalIgnoreCase)))
                    {
                        foreach (RfLib.InstrumentSetup.InstrumentTypeData.InstrumentTypeSetting instrumentsetup in path)
                        {
                            if (instrumentsetup.HighLimit < frequency || instrumentsetup.LowLimit > frequency)
                            {
                                var err = new RFLogError();
                                err.Type = ErrorType.FrequencyOutOfRange;
                                err.Level = EnumErrorLevel.Error;
                                err.Message = string.Format("UltraPac80 Source Frequency is 80MHz to 1kHz, WiSrc Pin Frequency is Out Of Range {0}.", log.Interface);
                                RFLogManager.Push(err, logName, row, "AUTOGEN");
                            }
                        }
                        break;
                    }
                }
            }
        }

        public static BenchLogFile GetLogPatternName(string originpath)
        {
            var loginfo = new BenchLogFile();
            //List<BenchLogItem> logs = new List<BenchLogItem>();
            loginfo.Offset = -1;
            string line = "";
            int i = 0;
            bool isNeedDelete = false;
            string tmp = originpath;
            while (IsFileLock(ref tmp))
            {
                if (!File.Exists(tmp))
                {
                    File.Copy(originpath, tmp);
                    isNeedDelete = true;
                }
                ;
            }
            var sr = new StreamReader(tmp);
            bool isFirst = true;
            int row = 0;
            while ((line = sr.ReadLine()) != null)
            {
                row++;
                try
                {
                    //OPERATION,INTERFACE,ADDRESS,MSB,LSB,REG_FIELD_NAME,DATA,DEFAULT,FUSE_NAME.
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }

                    if (loginfo.Offset == -1 && line.ToUpper().Contains("OPERATION"))
                    {
                        loginfo.Offset =
                            line.Split(',')
                                .ToList()
                                .FindIndex(p => p.Equals("OPERATION", StringComparison.OrdinalIgnoreCase));
                        //BenchLogItem.ResetIndex();
                        loginfo.GetHeader(line);
                    }
                    var log = new BenchLogItem(line, loginfo, row, false, ref isFirst);
                    if (log.Operation.Equals("PRINT", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(log.Operation))
                    {
                        continue;
                    }

                    if (Regex.IsMatch(log.Operation, "PATNAME", RegexOptions.IgnoreCase))
                    {
                        loginfo.PatternSubName = log.RawDatas[3];
                        loginfo.Type = log.RawDatas[2];
                        loginfo.Interface = log.RawDatas[1];
                        loginfo.Version = string.IsNullOrEmpty(log.RawDatas[4]) ? "1" : log.RawDatas[4];
                        loginfo.DateCode = string.IsNullOrEmpty(log.RawDatas[5]) ? "" : log.RawDatas[5];
                        continue;
                    }
                    if (Regex.IsMatch(log.Operation, @"SEQUENCE[\s_]*ORDER", RegexOptions.IgnoreCase))
                    {
                        loginfo.Inits.Add(log.Interface);
                        continue;
                    }
                    if (Regex.IsMatch(log.Operation, "OPERATION|print", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    ;
                }
                i++;
                if (i >= 5)
                {
                    break;
                }
            }
            sr.Close();
            if (isNeedDelete)
            {
                File.Delete(tmp);
            }

            return loginfo;
        }

        private static void CheckPatternName(int row, string log, string subblock, string version)
        {
            if (string.IsNullOrEmpty(subblock))
            {
                var err = new RFLogError();
                err.Type = ErrorType.PatNameRule;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("Missing \"PatName\" syntax on this bench log");
                RFLogManager.Push(err, log, row, "OPERATION");
            }
            else
            {
                if (subblock.Contains("_"))
                {
                    if (subblock.Count(p => p == '_') != 1)
                    {
                        var err = new RFLogError();
                        err.Type = ErrorType.PatNameRule;
                        err.Level = EnumErrorLevel.Error;
                        err.Message = string.Format("\"_\" should contains only one", subblock);
                        RFLogManager.Push(err, log, row, "MSB");
                    }
                    string[] sgmts = subblock.Split('_');
                    for (int i = 0; i < sgmts.Count(); i++)
                    {
                        string msg = "";
                        switch (i)
                        {
                            case 0:
                                if (sgmts[i].Length > 8)
                                {
                                    msg = string.Format("First segment {0}, characters exceed 8", sgmts[i]);
                                }

                                break;
                            case 1:
                                if (sgmts[i].Length > 8)
                                {
                                    msg = string.Format("Second segment {0}, characters exceed 8", sgmts[i]);
                                }

                                break;
                            default:
                                break;
                        }
                        if (!string.IsNullOrEmpty(msg))
                        {
                            var err = new RFLogError();
                            err.Type = ErrorType.PatNameRule;
                            err.Level = EnumErrorLevel.Error;
                            err.Message = msg;
                            RFLogManager.Push(err, log, row, "MSB");
                        }
                    }

                }
                else
                {
                    var err = new RFLogError();
                    err.Type = ErrorType.PatNameRule;
                    err.Level = EnumErrorLevel.Error;
                    err.Message = string.Format("{0} need contains '_', please check", subblock);
                    RFLogManager.Push(err, log, row, "MSB");
                }

            }
            if (!Regex.IsMatch(version, @"^[0-9]\d*$") && !string.IsNullOrEmpty(version))
            {
                var err = new RFLogError();
                err.Type = ErrorType.PatNameRule;
                err.Level = EnumErrorLevel.Error;
                err.Message = string.Format("Pattern version cannot be negative {0}", version); // Module Name can not be negative
                RFLogManager.Push(err, log, row, "Version");
            }

        }

        private Register CreateRegister(BenchLogItem log, string defValue, bool isNeedShift)
        {
            var result = new Register();
            string binStr;
            if (string.IsNullOrEmpty(defValue) || !defValue.StartsWith("0x"))
            {
                if (string.IsNullOrEmpty(defValue) && string.IsNullOrEmpty(log.FieldVal))
                {
                    binStr = ConvertValToBin("0x0", PatBit, false, isNeedShift);
                }
                else
                {
                    binStr = ConvertValToBin(log.FieldVal, PatBit, false, isNeedShift);
                }
            }
            else
            {
                binStr = ConvertValToBin(defValue, PatBit, false, isNeedShift);
            }

            result.Address = log.Address;
            result.Length = PatBit.ToString();
            result.Data = binStr;
            return result;
        }
    }

}




//WRITE,JTAG,0x4400c01c,4,4,bennutc_wl01.regs.bennutc.pd_radio.common_sys.XTAL.top_xtal_reg_7.top_xtal_dctp_en_1p0_da,0x00000001,,
[Serializable]
public class Register
{
    public Dictionary<string, string> SrcDic = new Dictionary<string, string>();
    public Dictionary<string, string> CapDic = new Dictionary<string, string>();
    public string Address;
    public string Length;
    //public List<Field> Fields = new List<Field>();
    public string Data = "";

    public bool IsNeedSouce
    {
        get { return Data.Any(p => p.Equals('D')); }
    }

    public List<string> GetFieldList(VectorInfo infoV)
    {
        //pd_radio__common_sys__TOP_MISC__TOP_MISC__top_misc_reg_6__top_5g_misc_spares2_1p0_da _14_to_10
        var nameList = new List<string>();
        var result = Enumerable.Range(0, 32).Select(p => string.Format("[{0}]", p)).ToList();
        foreach (KeyValuePair<string, string> src in SrcDic)
        {
            int msb = int.Parse(src.Key.Split('_')[1].Substring(1));
            int lsb = int.Parse(src.Key.Split('_')[2].Substring(1));
            for (int i = lsb; i <= msb; i++)
            {
                result[i] = string.Format("{0}[{1}]", src.Key, i - lsb);
            }
        }
        foreach (string cap in CapDic.Keys)
        {
            int msb = int.Parse(cap.Split('_')[1].Substring(1));
            int lsb = int.Parse(cap.Split('_')[2].Substring(1));
            for (int i = lsb; i <= msb; i++)
            {
                result[i] = string.Format("{0}[{1}]", cap, i - lsb);
            }
        }
        CapDic.Clear();

        return result;
    }
}
public class VectorInfo
{
    private double _time = 1000.0 / 24.0; //24MHz in nsec 41
    public double CurrentTime = 0;
    public int VectorCount = 0;
    public int RealCount = 0;
    public int WaitCount = 0;
    public string LastVector = "";
    private string _currentSourceField = "";
    public int CurrentSourceSgmt = -1;
    public int CurrentCaptureSgmt = -1;

    public void Update()
    {
        VectorCount++;
        if (WaitCount == 0)
        {
            RealCount++;
        }
        else
        {
            RealCount = RealCount + WaitCount;
        }

        WaitCount = 0;
        CurrentTime = Math.Round(_time * RealCount);
    }

    public void CheckSourceInfo(string regName)
    {
        if (regName.Split('[')[0] != _currentSourceField)
        {
            _currentSourceField = regName.Split('[')[0];
            CurrentSourceSgmt++;
        }
    }

    public void Clear()
    {
        _currentSourceField = "";
    }

}

public class NamingBox
{
    //vm_vector PP_BTCA0_S_PLLP_AN_RFD1_PFF_JTG_UNS_ALLFRV_SI_XTALLDO1_VOLTAGE_5_A0_2002191758
    //srm_vector PP_BTCA0_S_PLLP_AN_RFD1_PFF_JTG_UNS_ALLFRV_SI_XTALLDO1_VOLTAGE_srm_meas
    //global subr PP_BTCA0_S_PLLP_AN_RFD1_PFF_JTG_UNS_ALLFRV_SI_XTALLDO1_VOLTAGE_digsrc:
    //call PP_BTCA0_S_PLLP_AN_RFD1_PFF_JTG_UNS_ALLFRV_SI_XTALLDO1_VOLTAGE_digsrc 
    public string PatternFullName;
    public string PatternName;
    public string PatternSubr;
    public string SrmVecName;
    public static string PattPrefixed = "_S_PL00_AN_ARFX_PFF_JTG_UNS_ALLFRV_SI_";
    //public static string PattternPrefixedHT = "HT_PXMA0_S_PL00_AN_ARFX_PFF_JTG_UNS_ALLFRV_SI_";
    public NamingBox(BenchLogFile file, string patternName)
    {
        PatternName = patternName;
        PatternFullName = patternName;
        SrmVecName = file.Logs.Exists(p => Regex.IsMatch(p.Operation, "meas", RegexOptions.IgnoreCase) ||
            (Regex.IsMatch(p.Operation, "autogen", RegexOptions.IgnoreCase) && Regex.IsMatch(p.Operation, "TE_OPERATION", RegexOptions.IgnoreCase))
            )
            ? PatternName + "_srm_meas" : patternName + "_srm";
        var subrPostFixed = new List<string>();
        if (file.Logs.Exists(p => Regex.IsMatch(p.Operation, "read_cap", RegexOptions.IgnoreCase)))
        {
            subrPostFixed.Add("cap");
        }

        if (file.Logs.Exists(p => Regex.IsMatch(p.Operation, "write_src|write_trim", RegexOptions.IgnoreCase)))
        {
            subrPostFixed.Add("src");
        }

        if (subrPostFixed.Count != 0)
        {
            subrPostFixed.Insert(0, "dig");
        }
        else if (!string.IsNullOrEmpty(SrmVecName))
        {
            subrPostFixed.Add("meas");
        }

        if (subrPostFixed.Count != 0)
        {
            PatternSubr = string.Format("{0}_{1}", PatternName, string.Join("", subrPostFixed));
        }
    }

}

public class MappingItem
{
    public string Log;
    public string Pattern;
    public string Silicon;
    public string Date;
    public string Version;
    public List<string> Inits = new List<string>();
    public string GetFullPattern()
    {
        return string.Format("{0}_{1}_{2}_{3}", Pattern, Version, Silicon, Date);
    }
}

public class PythonFunctionParser
{
    private static readonly Regex _defPattern = new Regex(
        @"(?m)^\s*(?:@.*\r?\n)*def\s+(?<_name>[A-Za-z_]\w*)\s*\((?<_params>[^)]*)\)",
        RegexOptions.Compiled);

    public static Dictionary<string, List<string>> ParseFromFile(string filePath)
    {
        string source = filePath;
        if (File.Exists(filePath))
        {
            source = File.ReadAllText(filePath);
        }
        return ParseFromString(source);
    }

    public static Dictionary<string, List<string>> ParseFromString(string source)
    {
        var map = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (Match m in _defPattern.Matches(source))
        {
            string funcName = m.Groups["_name"].Value;
            string rawParams = m.Groups["_params"].Value;

            List<string> parameters = ExtractParamNames(rawParams);

            map.Add(funcName, parameters);
        }

        return map;
    }

    private static List<string> ExtractParamNames(string raw)
    {
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return list;
        }

        List<string> parts = SafeSplit(raw, ',');

        foreach (string part in parts)
        {
            string p = part.Trim();
            if (p.Length == 0)
            {
                continue;
            }

            int colon = IndexOfTopLevel(p, ':');
            if (colon >= 0)
            {
                p = p.Substring(0, colon).Trim();
            }

            int eq = IndexOfTopLevel(p, '=');
            if (eq >= 0)
            {
                p = p.Substring(0, eq).Trim();
            }

            p = p.Replace(" ", "");

            if (p.Length > 0)
            {
                list.Add(p);
            }
        }

        return list;
    }

    private static int IndexOfTopLevel(string s, char ch)
    {
        int par = 0, br = 0, bc = 0;
        bool inSingle = false, inDouble = false;

        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];

            if ((inSingle || inDouble) && c == '\\')
            {
                i++;
                continue;
            }

            if (!inDouble && c == '\'')
            {
                inSingle = !inSingle;
            }
            else if (!inSingle && c == '"')
            {
                inDouble = !inDouble;
            }
            else if (!inSingle && !inDouble)
            {
                if (c == '(')
                {
                    par++;
                }
                else if (c == ')')
                {
                    par--;
                }
                else if (c == '[')
                {
                    br++;
                }
                else if (c == ']')
                {
                    br--;
                }
                else if (c == '{')
                {
                    bc++;
                }
                else if (c == '}')
                {
                    bc--;
                }
                else if (c == ch && par == 0 && br == 0 && bc == 0)
                {
                    return i;
                }
            }
        }
        return -1;
    }

    private static List<string> SafeSplit(string s, char sep)
    {
        var list = new List<string>();
        var sb = new StringBuilder();

        int par = 0, br = 0, bc = 0;
        bool inSingle = false, inDouble = false;

        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];

            if ((inSingle || inDouble) && c == '\\')
            {
                sb.Append(c);
                if (i + 1 < s.Length)
                {
                    sb.Append(s[++i]);
                }

                continue;
            }

            if (!inDouble && c == '\'')
            {
                inSingle = !inSingle;
                sb.Append(c);
            }
            else if (!inSingle && c == '"')
            {
                inDouble = !inDouble;
                sb.Append(c);
            }
            else if (!inSingle && !inDouble)
            {
                if (c == '(')
                { par++; sb.Append(c); }
                else if (c == ')')
                { par--; sb.Append(c); }
                else if (c == '[')
                { br++; sb.Append(c); }
                else if (c == ']')
                { br--; sb.Append(c); }
                else if (c == '{')
                { bc++; sb.Append(c); }
                else if (c == '}')
                { bc--; sb.Append(c); }
                else if (c == sep && par == 0 && br == 0 && bc == 0)
                {
                    list.Add(sb.ToString());
                    sb.Length = 0;
                }
                else
                {
                    sb.Append(c);
                }
            }
            else
            {
                sb.Append(c);
            }
        }

        if (sb.Length > 0)
        {
            list.Add(sb.ToString());
        }

        return list;
    }
}
