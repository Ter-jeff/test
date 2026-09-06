using System.Collections.Concurrent;
using System.ComponentModel;
using System.Text.RegularExpressions;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using NLog;

using OfficeOpenXml;

using RF_PatternTool.PatternGen;
using RF_PatternTool.PatternStruct;

using RfLib.Dvdc.Reader.CapturePostProcess;

using RFPatternTool;

using TestPlanLib.Static;


namespace RF_PatternTool
{
    public enum GenerateType
    {
        ARF, HSC, HTOL
    }
    public class PatternResult
    {
        public string OutputDir { get; set; }

        public List<PatternFile> Files { get; set; } = new List<PatternFile>();

        public bool NeedGlobalSub { get; set; }

        public Dictionary<string, string> WriteSrcRows { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
    public class PatternFile
    {
        public string FileName { get; set; }
        public List<string> Content { get; set; } = new List<string>();
    }
    public class MeasType
    {
        public static readonly string MeasV = "MeasV";
        public static readonly string MeasI = "MeasI";
        public static readonly string MeasR = "MeasR";
        public static readonly string MeasF = "MeasF";
        public static readonly string WiMeas = "WiMeas";
        public static readonly string WiSrc = "WiSrc";
        public static readonly string LXSrc = "LXSrc";
        public static readonly string LxRx = "LxRx";
    }

    public class WriteTrimInfo : IComparable<WriteTrimInfo>, IEquatable<WriteTrimInfo>
    {
        public string IndexAddress = "";
        public int IndexMSB = -1;
        public int IndexLSB = -1;
        public int IndexFIELD_VAL = -1;

        public WriteTrimInfo(string address, int msb, int lsb, int field_val)
        {
            IndexAddress = address;
            IndexMSB = msb;
            IndexLSB = lsb;
            IndexFIELD_VAL = field_val;
        }

        public int CompareTo(WriteTrimInfo wt)
        {
            if (IndexAddress == wt.IndexAddress && IndexMSB == wt.IndexMSB && IndexLSB == wt.IndexLSB)
            {
                return 0;
            }

            return 1;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as WriteTrimInfo);
        }

        public bool Equals(WriteTrimInfo other)
        {
            return !(other is null) &&
                   IndexAddress == other.IndexAddress &&
                   IndexMSB == other.IndexMSB &&
                   IndexLSB == other.IndexLSB;
        }

        public override int GetHashCode()
        {
            int hashCode = -1553780428;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(IndexAddress);
            hashCode = hashCode * -1521134295 + IndexMSB.GetHashCode();
            hashCode = hashCode * -1521134295 + IndexLSB.GetHashCode();
            return hashCode;
        }
    }

    public class PatternGenBusiness
    {
        private readonly BodyPattern _bodyP = new BodyPattern();
        private readonly ReadPattern _readP = null;
        private ReadPattern _read64P = null;
        private SubRoutine _subrP = new SubRoutine();
        private WritePattern _writeP = null;
        private WritePattern _write64P = null;
        private MatchLoopPattern _matchP = null;
        public GenerateType GenType = GenerateType.ARF;
        private List<string> _rEuseDictionary = new List<string>();
        private static string _project = "";
        public BindingList<PatternItem> Patterns = new BindingList<PatternItem>();
        public Dictionary<string, List<string>> MatchLoopInfo = new Dictionary<string, List<string>>();
        public ConcurrentDictionary<string, byte> pairSet = new ConcurrentDictionary<string, byte>();
        public List<PostProcessSheetRow> PostProcessRows = new List<PostProcessSheetRow>();
        private List<PatternGenItem> _itemInfos = new List<PatternGenItem>();
        private PatternGenItem _itemInfo = new PatternGenItem();
        private Dictionary<string, List<string>> _usedDic = new Dictionary<string, List<string>>();
        public static bool IsCheck = false;
        public List<List<string>> EfusebitDefRows = new List<List<string>>();
        public static string TemplateSCGH = "";
        private double _freq = 0;
        private HashSet<WriteTrimInfo> _wtrecord = new HashSet<WriteTrimInfo>();
        private bool _isR16Proj = false;
        private HashSet<string> _fullAddrFor64InBin = new HashSet<string>();
        public int PatBit = 32;
        private static int _threadFlag = 0;

        private StreamWriter _swlog = null;
        public static string CurrDateCode;
        private HashSet<string> _allPats = new HashSet<string>();
        public bool IsAllowOverWrite = true;
        public static bool IsAddComment = true;

        private Regex _regtest = new Regex("meas|INJECT|autogen", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private readonly string _reg_meas = @"//TE_OPERATION\s+/*/*(?<InstType>\w+)\s+(?:Pin|DisconnectPins)\s*=\s*(?<PinName>[\w,:-]+)";
        private readonly string _reg_loopback = @"//TE_OPERATION\s*//Loopback\s+(?<type>pin|path|DisconnectPins)=(?<loopbackpath>[^/]+)(?:\/\/expected_attenuation=(?<expectedattenuation>.*))?";

        public PatternGenBusiness()
        {
        }

        public PatternGenBusiness(List<PatternGenItem> genItem)
        {
            _itemInfos = genItem;
        }

        public PatternGenBusiness(string registermap, ref List<PatternGenItem> genItem, GenerateType type, double jtagFreq,
            bool isOverWrite, bool isAddComment)
        {
            GenType = type;
            _itemInfos = genItem;
            _readP = new ReadPattern(jtagFreq);
            _read64P = new ReadPattern(jtagFreq);
            _writeP = new WritePattern(jtagFreq);
            _write64P = new WritePattern(jtagFreq);
            _matchP = new MatchLoopPattern(jtagFreq);
            _freq = jtagFreq;
            IsAllowOverWrite = isOverWrite;
            IsAddComment = isAddComment;
        }

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

        public static Dictionary<string, Register> LoadRegisters(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return new Dictionary<string, Register>();
            }

            var sr = new StreamReader(filePath);
            var registers = new Dictionary<string, Register>();
            string para = "";
            Register currentRegister = null;
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                try
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }

                    if (line.Contains('$'))
                    {
                        para = "";
                    }

                    para = para + line;

                    if (line.Contains(';'))
                    {
                        string reg_key = @"\$(?<key>\w+)";
                        string key = Regex.Match(para, reg_key, RegexOptions.IgnoreCase).Groups["key"].Value;
                        if (key.Equals("HIER", StringComparison.OrdinalIgnoreCase))
                        {
                            currentRegister = new Register();
                        }
                        else if (key.Equals("APB_ADDR", StringComparison.OrdinalIgnoreCase))
                        {
                            currentRegister.Address = para.Split('=')[1].Trim(';').Trim();
                        }
                        else if (key.Equals("CHAIN_LEN", StringComparison.OrdinalIgnoreCase))
                        {
                            currentRegister.Length = para.Split('=')[1].Trim(')').Trim(';').Trim();
                        }
                        else if (key.Equals("type", StringComparison.OrdinalIgnoreCase))
                        {
                            List<string> fields = GetRegisters(para.Split('=')[1]);
                            for (int i = fields.Count - 1; i >= 0; i--)
                            {
                                currentRegister.Data = fields[i] + currentRegister.Data;
                            }
                            currentRegister.Data.Reverse();

                            registers.Add(currentRegister.Address, currentRegister);
                            currentRegister = null;
                        }
                    }
                }
                catch (Exception)
                {
                    ;
                }
            }
            return registers;
        }

        public static MappingItem GetNamingInfo(string filename, GenerateType type, string projName, string siliconVer, bool isFullSweep, bool isDebugMode = false)
        {
            var patternName = new MappingItem();
            BenchLogFile shortinfo = GetLogPatternName(filename, isFullSweep);
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
            if (type == GenerateType.HTOL)
            {
                patternName.Pattern = "HT_" + projName + siliconVer + NamingBox.PattPrefixed + logname.Split('#')[0];
            }

            if (type == GenerateType.ARF)
            {
                patternName.Pattern = shortinfo.Interface + "_" + projName + siliconVer + NamingBox.PattPrefixed + logname.Split('#')[0];
            }

            patternName.Pattern = patternName.Pattern.Replace("ARFX", shortinfo.Type);
            patternName.Silicon = siliconVer;
            patternName.Date = isDebugMode ? "DEBUGMODE" : string.IsNullOrEmpty(shortinfo.DateCode) ? CurrDateCode : shortinfo.DateCode;
            patternName.Version = shortinfo.Version;
            patternName.Inits.AddRange(shortinfo.Inits);
            return patternName;

        }

        public static MappingItem GetNamingInfo(BenchLogFile file, GenerateType type, string projName, string siliconVer)
        {
            var patternName = new MappingItem();
            string logname = file.PatternSubName;

            if (logname.Contains('#'))
            {
                patternName.Version = logname.Split('#')[1];
            }
            else
            {
                patternName.Version = "1";
            }

            patternName.Log = logname.Split('#')[0];
            if (type == GenerateType.HTOL)
            {
                patternName.Pattern = "HT_" + projName + siliconVer + NamingBox.PattPrefixed + logname.Split('#')[0];
            }

            if (type == GenerateType.ARF)
            {
                patternName.Pattern = file.Interface + "_" + projName + siliconVer + NamingBox.PattPrefixed + logname.Split('#')[0];
            }

            patternName.Pattern = patternName.Pattern.Replace("ARFX", file.Type);
            patternName.Silicon = siliconVer;
            patternName.Date = string.IsNullOrEmpty(file.DateCode) ? CurrDateCode : file.DateCode;
            patternName.Version = file.Version;
            return patternName;

        }

        public PatternResult PrintPatternAtp(string filepath, MappingItem patternname,
            Dictionary<string, Register> tmpregisters, Logger logger, string outputDir,
            string projName, string siliconVer, string preSetupPats, HashSet<string> pinMapPins,
            bool isR16K, bool isFullSweep, HashSet<string> addrFor64InBin, bool isDebugMode = false)
        {
            _isR16Proj = isR16K;
            _fullAddrFor64InBin = addrFor64InBin;

            var result = new PatternResult
            {
                OutputDir = outputDir,
                Files = new List<PatternFile>()
            };

            logger.Info($"Generate {patternname.Log} ...");
            string logDirect = Path.GetDirectoryName(filepath);
            Dictionary<string, string> localWriteSrc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _rEuseDictionary.Clear();

            #region Read Log
            List<BenchLogFile> benchLogs = ReadLogFile(filepath, isFullSweep, isDebugMode, outputDir);
            CollectWriteSrcRows(benchLogs, localWriteSrc);
            result.WriteSrcRows = localWriteSrc;
            bool isFirstPat = true;
            bool isFW = benchLogs.Exists(p => p.Type.Equals("FARF", StringComparison.OrdinalIgnoreCase)) ||
                (benchLogs.SelectMany(p => p.Logs).ToList().Exists(p => Regex.IsMatch(p.Operation, "Wait_For", RegexOptions.IgnoreCase)));
            #endregion

            string logreport = string.Format("ModRegLog_{0}_{1}.csv", Path.GetFileNameWithoutExtension(filepath), DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            _swlog = new StreamWriter(Path.Combine(outputDir, logreport));
            List<string> inits = preSetupPats.Split(',').ToList();

            Dictionary<string, Register> registers = new Dictionary<string, Register>(StringComparer.OrdinalIgnoreCase);

            Dictionary<string, long> checkTime = new Dictionary<string, long>();
            bool measCallExisted = false;
            foreach (BenchLogFile benchLog in benchLogs)
            {
                CapTrimInfo readCapTrim = null;
                _usedDic.Clear();
                PostProcessRows.Clear();
                string logname = string.IsNullOrEmpty(benchLog.PatternSubName)
                    ? Path.GetFileNameWithoutExtension(filepath)
                    : benchLog.PatternSubName;

                if (_isR16Proj)
                {
                    registers = FilterBanks(tmpregisters,
                        benchLog.Logs.Where(p => !string.IsNullOrEmpty(p.Address)).Select(p => p.Address).Distinct().ToList());
                }

                var info_address = benchLog.Logs.Where(p => p.Operation.IndexOf("write", StringComparison.CurrentCultureIgnoreCase) >= 0).Select(p => p.Address).Distinct().ToList();

                bool isSrcBefore = registers
                    .Where(reg => info_address.Contains(reg.Key))
                    .Any(reg => reg.Value.IsNeedSouce);

                #region Generate Pattern

                //A. operation sequence => 
                // 1. register value should flush with table values
                // 2. sequence_order -> update  related pattern log
                //B. write pattern =>
                // 1. write -> read log address and data -> update inherit data at current address -> print out fixed value  (wait time repeat 52 or 103?)
                // 2. measV/I/R/F -> print call subroutine(if exist -> create subroutine below) -> need add pattern comment
                // 3. write_src -> same as "write" -> but need print out "D" at specific bit position
                // 4. wait add wait time with related vector (assume JTAG frequency with 24MHz, )
                //
                var infoV = new VectorInfo();
                bool isNeedBreak = false;
                MappingItem patternName = GetNamingInfo(benchLog, GenType, projName, siliconVer);
                _swlog.WriteLine("Create item : {0}_{1}_{2}", benchLog.PatternSubName, benchLog.Version, string.IsNullOrEmpty(benchLog.DateCode) ? CurrDateCode : benchLog.DateCode);
                string block = isFW ? "FW" : "ARF";
                string cppsetup = string.Format("{1}_{0}_DSSCSetup_Post_Process", patternName.Log, block);
                _itemInfo = new PatternGenItem();
                _itemInfo.InitDictionary.AddRange(inits);
                inits.Clear();
                _itemInfos.Add(_itemInfo);

                string toolAtp = patternName.GetFullPattern() + ".atp";
                var patFile = new PatternFile
                {
                    FileName = toolAtp
                };
                result.Files.Add(patFile);
                _itemInfo.Pattern = patternName.Pattern;
                _itemInfo.ScghName = new KeyValuePair<string, string>(logname, patternName.Pattern);
                if (_allPats.Contains(patternName.Pattern))
                {
                    continue;
                }
                else
                {
                    _allPats.Add(patternName.Pattern);
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                List<string> patrows = new List<string>();

                bool isSource = false;
                bool isCap =
                    benchLog.Logs.Exists(p =>
                        Regex.IsMatch(p.Operation, "Read_Cap", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(p.Operation, "Spac_Mem_Read", RegexOptions.IgnoreCase));
                if (isFirstPat)
                {
                    InitSetup();
                    isFirstPat = false;
                }

                string lastVecFromPreBody = GetLastVec();
                var nameBase = new NamingBox(benchLog, patternName.Pattern);

                _itemInfo.Patset = new KeyValuePair<string, string>(patternName.Pattern, patternName.GetFullPattern());

                List<BenchLogItem> trimLogs = new List<BenchLogItem>();
                List<BenchLogItem> usedTrimLogs = new List<BenchLogItem>();
                bool isTrimStart = false;
                bool isMeasCall = false;

                #region generate ATP with Bench log

                string trimName = benchLog.Type + "_" + benchLog.PatternSubName;
                string forceinfo = "";
                string trimInfo = "";
                string trimMeasName = "";
                bool isIncludeSubr = false;
                bool readCapTrimFlag = false;
                bool loopBackFlag = false;

                HashSet<string> matchLoopSAddressSRM = new HashSet<string>();

                _readP.PatSubName = benchLog.PatternSubName;
                _readP.StoreIndex = 0;
                _read64P.PatSubName = benchLog.PatternSubName;
                _read64P.StoreIndex = 0;

                foreach (BenchLogItem log in benchLog.Logs)
                {
                    try
                    {
                        if (isNeedBreak)
                        {
                            if (log.Operation.Equals("trim_end", StringComparison.OrdinalIgnoreCase))
                            {
                                isNeedBreak = false;
                                _wtrecord.Clear();
                            }
                            continue;
                        }

                        BenchLogItem usedLog = usedTrimLogs.FirstOrDefault(p => p.Address == log.Address && p.LSB == log.LSB && p.MSB == log.MSB);

                        bool isSkipLog = DispatchBenchLogOperation(patrows, benchLog, log, usedLog, registers,
                            tmpregisters, infoV, logger, logDirect, patternName, projName, siliconVer, isFullSweep,
                            isDebugMode, pinMapPins, nameBase, lastVecFromPreBody, isCap, matchLoopSAddressSRM,
                            trimLogs, usedTrimLogs, ref forceinfo, ref isNeedBreak, ref isTrimStart, ref isSource,
                            ref isMeasCall, ref measCallExisted, ref isIncludeSubr, ref readCapTrimFlag,
                            ref loopBackFlag, ref trimInfo, ref trimMeasName, ref readCapTrim);
                        if (isSkipLog)
                        {
                            continue;
                        }

                        if (isNeedBreak)
                        {
                            if (!isIncludeSubr)
                            {
                                string comment = string.Format(@"//TE_OPERATION MeasN Pin = ");
                                PrintCallSubroutine(patrows, infoV, nameBase.PatternSubr, lastVecFromPreBody, trimInfo, comment);
                                isMeasCall = true;
                                measCallExisted = true;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message);
                    }
                }
                AppendMiscInfoRows(patrows, trimMeasName, block, cppsetup, patternName, isFW, readCapTrimFlag,
                    loopBackFlag);

                patrows.Add("}");
                #endregion

                var headerLines = new List<string>();
                using (var swMem = new StringWriter())
                {
                    PrintswBody(swMem, infoV, patternName.GetFullPattern(), isSource, isCap, isMeasCall, matchLoopSAddressSRM);

                    headerLines.AddRange(swMem.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None));
                }
                patFile.Content.AddRange(headerLines);
                patFile.Content.AddRange(patrows);

                _itemInfo.Cpp.AddRange(PostProcessRows);
                #endregion

            }
            if (measCallExisted && (Interlocked.CompareExchange(ref _threadFlag, 1, 0) == 0))
            {
                result.NeedGlobalSub = true;
                List<string> teopeartionSRM = new List<string>();
                CreateSubRoutine(teopeartionSRM, "Global_TEOPEARTION_SRM", "Global_TEOPEARTION_meas");

                var patFile = new PatternFile
                {
                    FileName = "Global_TEOPEARTION.atp"
                };

                patFile.Content.AddRange(teopeartionSRM);

                result.Files.Add(patFile);

                _itemInfo.PatSubr.Add(new KeyValuePair<string, string>("Global_TEOPEARTION", "Global_TEOPEARTION_SRM"));
            }
            AppendMatchLoopFiles(result);
            _swlog.Close();
            logger.Info($"Done {patternname.Log}");
            return result;
        }

        private static void CollectWriteSrcRows(List<BenchLogFile> benchLogs, Dictionary<string, string> localWriteSrc)
        {
            foreach (BenchLogFile log in benchLogs)
            {
                foreach (KeyValuePair<string, string> kv in log.WriteSrcRows)
                {
                    localWriteSrc[kv.Key] = kv.Value;
                }
            }
        }

        private bool DispatchBenchLogOperation(List<string> patrows, BenchLogFile benchLog, BenchLogItem log,
            BenchLogItem usedLog, Dictionary<string, Register> registers, Dictionary<string, Register> tmpregisters,
            VectorInfo infoV, Logger logger, string logDirect, MappingItem patternName, string projName,
            string siliconVer, bool isFullSweep, bool isDebugMode, HashSet<string> pinMapPins, NamingBox nameBase,
            string lastVecFromPreBody, bool isCap, HashSet<string> matchLoopSAddressSRM, List<BenchLogItem> trimLogs,
            List<BenchLogItem> usedTrimLogs, ref string forceinfo, ref bool isNeedBreak, ref bool isTrimStart,
            ref bool isSource, ref bool isMeasCall, ref bool measCallExisted, ref bool isIncludeSubr,
            ref bool readCapTrimFlag, ref bool loopBackFlag, ref string trimInfo, ref string trimMeasName,
            ref CapTrimInfo readCapTrim)
        {
            switch (log.Operation.ToLower())
            {
                case "sequence order":
                case "sequence_order":
                    UpdateTableValue(logDirect, log.Interface, tmpregisters, patternName.Pattern, projName, siliconVer, isFullSweep, isDebugMode);
                    break;
                case "reg_write":
                    PrintWrite(patrows, log, registers, infoV, false, ref isSource, true);
                    break;
                case "write_trim":
                    PrintWriteTrimLog(patrows, benchLog, log, registers, infoV, trimLogs, usedTrimLogs,
                        readCapTrim, ref isNeedBreak, ref isTrimStart, ref isSource);
                    break;
                case "write":
                    PrintWriteOperationLog(patrows, log, usedLog, registers, infoV, trimLogs, usedTrimLogs,
                        ref isNeedBreak, ref isTrimStart, ref isSource);
                    break;
                case "write_src":
                    PrintWriteSrcLog(patrows, log, registers, infoV, ref isSource);
                    break;
                case "reg_read":
                    break;
                case "otp_write":
                    patrows.Add(string.Format("//  TE_SETUP MiscInfo = SetEfuse : {0} = {1}", log.Address, log.Interface));
                    break;
                case "read":
                    break;
                case "read_compare":
                    PrintRead(patrows, log, registers, infoV, isTrimStart, _itemInfo);
                    break;
                case "wait_for":
                    PrintWaitForLog(patrows, log, registers, infoV, logger, matchLoopSAddressSRM,
                        lastVecFromPreBody, trimInfo);
                    break;
                case "spac_mem_read":
                    PrintSpacMemReadLog(patrows, log, registers, infoV, isTrimStart);
                    break;
                case "read_cap":
                    PrintReadCapLog(patrows, log, registers, infoV, isTrimStart, readCapTrim, ref trimMeasName);
                    break;
                case "wait":
                    _swlog.WriteLine("Wait,{0}", log.Interface);
                    PrintWait(patrows, infoV, log.Interface, lastVecFromPreBody);
                    break;
                case "lut":
                    if (_itemInfo.LutItem == null)
                    {
                        _itemInfo.LutItem = new LutItem(log.Address);
                    }
                    break;
                case "calc":
                    PrintCalcLog(log, isCap, isTrimStart, isFullSweep, readCapTrim);
                    break;
                case "forcev":
                case "forcei":
                    forceinfo = string.Format(@"//TE_OPERATION {0} Pin = {1} //{0} = {2} ", log.Operation, log.Interface, log.Address.Replace("0x", ""));
                    PrintCallSubroutine(patrows, infoV, nameBase.PatternSubr, lastVecFromPreBody, trimInfo, forceinfo);
                    forceinfo = "";
                    break;
                case "trim_end":
                    PrintTrimEndLog(patrows, benchLog, ref readCapTrim, ref readCapTrimFlag);
                    break;
                default:
                    bool isSkipLog = PrintTeOperationLog(patrows, benchLog, log, infoV, pinMapPins,
                        lastVecFromPreBody, trimLogs, isTrimStart, ref isIncludeSubr, ref isMeasCall,
                        ref measCallExisted, ref loopBackFlag, ref trimInfo, ref trimMeasName, ref readCapTrim);
                    if (isSkipLog)
                    {
                        return true;
                    }
                    break;
            }

            return false;
        }

        private void PrintWriteTrimLog(List<string> patrows, BenchLogFile benchLog, BenchLogItem log,
            Dictionary<string, Register> registers, VectorInfo infoV, List<BenchLogItem> trimLogs,
            List<BenchLogItem> usedTrimLogs, CapTrimInfo readCapTrim, ref bool isNeedBreak, ref bool isTrimStart,
            ref bool isSource)
        {
            if (!_isR16Proj)
            {
                WriteTrimInfo wt = new WriteTrimInfo(log.Address, log.MSB, log.LSB, int.Parse(log.FieldVal));
                if (_wtrecord.Contains(wt) && !benchLog.IsReadCapTrim)
                {
                    isNeedBreak = true;
                    return;
                }
                else if (readCapTrim != null)
                {
                    readCapTrim.SetTrim(log);
                    log.FuseName = "";
                    readCapTrim.AllCalcStoreNames.Add(new List<string>());
                }
                _wtrecord.Add(wt);
            }

            trimLogs.Add(log);
            usedTrimLogs.Add(log);
            if (GenType == GenerateType.HTOL)
            {
                log.FuseName = "";
            }

            log.Default = log.FuseName;
            log.Operation = "TrimCode";
            if (!_isR16Proj)
            {
                PrintWrite(patrows, log, registers, infoV, true, ref isSource);
            }
            isTrimStart = true;
        }

        private void PrintWriteOperationLog(List<string> patrows, BenchLogItem log, BenchLogItem usedLog,
            Dictionary<string, Register> registers, VectorInfo infoV, List<BenchLogItem> trimLogs,
            List<BenchLogItem> usedTrimLogs, ref bool isNeedBreak, ref bool isTrimStart, ref bool isSource)
        {
            if (isTrimStart &&
                usedLog == null &&
                (trimLogs.FirstOrDefault(p => p.Address == log.Address && p.LSB == log.LSB && p.MSB == log.MSB) != null))
            {
                if (usedTrimLogs.Count == 0 && trimLogs.Count != 0)
                {
                    isNeedBreak = true;
                }
            }
            else if ((usedLog != null && _isR16Proj) && (usedLog.Address == log.Address && usedLog.LSB == log.LSB && usedLog.MSB == log.MSB))
            {
                log.Default = usedLog.FuseName;
                log.FuseName = usedLog.FuseName;
                log.RegFieldName = usedLog.RegFieldName;
                log.Operation = "TrimCode";
                PrintWrite(patrows, log, registers, infoV, true, ref isSource);
                usedTrimLogs.Remove(usedLog);
                isTrimStart = true;
            }
            else if (trimLogs.Count == 0 || !isNeedBreak)
            {
                PrintWrite(patrows, log, registers, infoV, false, ref isSource);
            }
        }

        private void PrintWriteSrcLog(List<string> patrows, BenchLogItem log, Dictionary<string, Register> registers,
            VectorInfo infoV, ref bool isSource)
        {
            if ((log.Default.StartsWith("0x", StringComparison.OrdinalIgnoreCase) || int.TryParse(log.Default, out int _) || string.IsNullOrEmpty(log.Default)) &&
                !string.IsNullOrEmpty(log.FuseName))
            {
                if (GenType == GenerateType.ARF)
                {
                    log.Default = log.FuseName;
                }
                else
                {
                    log.FuseName = "";
                }
            }
            PrintWrite(patrows, log, registers, infoV, false, ref isSource);
        }

        private void PrintWaitForLog(List<string> patrows, BenchLogItem log, Dictionary<string, Register> registers,
            VectorInfo infoV, Logger logger, HashSet<string> matchLoopSAddressSRM, string lastVecFromPreBody,
            string trimInfo)
        {
            string matchloopname = $"Global_{log.Address}_M{log.MSB}_L{log.LSB}_Val{log.FieldVal}";
            string pairkey = $"{log.Address}_{log.FieldVal}";
            matchLoopSAddressSRM.Add(matchloopname);
            bool isNew = pairSet.TryAdd(pairkey, 0);

            if (!MatchLoopInfo.ContainsKey(matchloopname) && isNew)
            {
                MatchLoopInfo.Add(matchloopname, PrintWaitFor(log, registers, infoV, logger));
            }

            PrintCallSubroutine(patrows, infoV, matchloopname, lastVecFromPreBody, trimInfo, "");
        }

        private void PrintSpacMemReadLog(List<string> patrows, BenchLogItem log,
            Dictionary<string, Register> registers, VectorInfo infoV, bool isTrimStart)
        {
            uint startAdd = Convert.ToUInt32(log.Address, 16);
            uint runtimes = Convert.ToUInt32(log.RegData, 10);

            var tmpPostProcRows = new List<PostProcessSheetRow>(PostProcessRows);
            PostProcessRows.Clear();

            for (int i = 0; i < runtimes; i++)
            {
                string newAddress = $"0x{startAdd + i * 4:X}";
                log.Address = newAddress;
                log.Capinfos.First().Address = newAddress;
                PrintRead(patrows, log, registers, infoV, isTrimStart, _itemInfo);
            }

            PostProcessSheetRow firstPostProcRows = PostProcessRows.First();
            uint firstbw = Convert.ToUInt32(firstPostProcRows.BitWidth, 10) * runtimes;
            firstPostProcRows.BitWidth = firstbw.ToString();
            PostProcessRows = new List<PostProcessSheetRow>() { firstPostProcRows };
            PostProcessRows.InsertRange(0, tmpPostProcRows);
        }

        private void PrintReadCapLog(List<string> patrows, BenchLogItem log, Dictionary<string, Register> registers,
            VectorInfo infoV, bool isTrimStart, CapTrimInfo readCapTrim, ref string trimMeasName)
        {
            PrintRead(patrows, log, registers, infoV, isTrimStart, _itemInfo);
            if (isTrimStart)
            {
                if (!string.IsNullOrEmpty(log.TestName))
                {
                    trimMeasName = log.TestName;
                }
            }
            if (readCapTrim != null)
            {
                readCapTrim.AllCalcStoreNames.Last().AddRange(PostProcessRows.Last().PostCalcs.Where(p => !string.IsNullOrEmpty(p.CalcStoreName)).Select(p => p.CalcStoreName));
            }
        }

        private void PrintCalcLog(BenchLogItem log, bool isCap, bool isTrimStart, bool isFullSweep,
            CapTrimInfo readCapTrim)
        {
            if (!isCap)
            {
                PostProcessRows.Add(new PostProcessSheetRow());
            }

            PostProcessSheetRow postprocess = PostProcessRows.Last();
            postprocess.AnalyzeCalc(log.PostProcess, log.TestName, log.LoLimit, log.HighLimit, log.Units, isTrimStart, isFullSweep);
            if (readCapTrim != null && !string.IsNullOrEmpty(postprocess.PostCalcs.Last().CalcStoreName))
            {
                readCapTrim.AllCalcStoreNames.Last().Add(postprocess.PostCalcs.Last().CalcStoreName);
            }
        }

        private void PrintTrimEndLog(List<string> patrows, BenchLogFile benchLog, ref CapTrimInfo readCapTrim,
            ref bool readCapTrimFlag)
        {
            if (readCapTrim != null)
            {
                patrows.Add(string.Format("//  TE_SETUP MiscInfo = Trimbits:{0};", string.Join(",", readCapTrim.TrimBits.Values)));
                foreach (string item in readCapTrim.DataInfo.Keys)
                {
                    patrows.Add(string.Format("//  TE_SETUP TrimFuseName = {0}", item));
                }
                PostProcessSheetRow postproc = PostProcessRows.Last();
                var postcalc = new PostCalcInfo { CalcEquation = readCapTrim.GetTrimCalcInfo(), HiLimit = readCapTrim.HighLimit, LowLimit = readCapTrim.LowLimit };
                postproc.PostCalcs.Add(postcalc);
                benchLog.IsReadCapTrim = false;
                readCapTrim = null;
                readCapTrimFlag = true;
                _wtrecord.Clear();
            }
        }

        private bool PrintTeOperationLog(List<string> patrows, BenchLogFile benchLog, BenchLogItem log,
            VectorInfo infoV, HashSet<string> pinMapPins, string lastVecFromPreBody, List<BenchLogItem> trimLogs,
            bool isTrimStart, ref bool isIncludeSubr, ref bool isMeasCall, ref bool measCallExisted,
            ref bool loopBackFlag, ref string trimInfo, ref string trimMeasName, ref CapTrimInfo readCapTrim)
        {
            if (GenType != GenerateType.ARF)
            {
                return true;
            }
            if (_regtest.IsMatch(log.Operation))
            {
                if (isTrimStart)
                {
                    isIncludeSubr = true;
                }

                string comment = log.GetSpecialData();

                if (Regex.IsMatch(log.Interface, @"TE_SETUP", RegexOptions.IgnoreCase))
                {
                    string reg_setup = @"//\s*TE_SETUP\s*(?<msg>.*)";
                    string setup =
                        Regex.Match(comment, reg_setup, RegexOptions.IgnoreCase).Groups["msg"]
                            .Value;
                    patrows.Add(string.Format("// TE_SETUP {0}", setup));
                    comment = Regex.Replace(log.Interface, reg_setup, "", RegexOptions.IgnoreCase);
                }

                if (log.PostProcess.StartsWith(@"StoreName:"))
                {
                    if (Regex.IsMatch(log.Operation, "autogen", RegexOptions.IgnoreCase))
                    {
                        string storeName = log.PostProcess.Split(':')[1];
                        comment = comment + "//MeasStoreName=" + storeName;
                    }
                }

                if (string.IsNullOrEmpty(comment.Trim('\"').Trim(',')))
                {
                    return true;
                }

                if (!Regex.IsMatch(comment, "TE_OPERATION", RegexOptions.IgnoreCase))
                {
                    if (Regex.IsMatch(log.Operation, "INJECT", RegexOptions.IgnoreCase))
                    {
                        comment = string.Format(@"//TE_OPERATION WiSrc Pin = {0}", log.Interface);
                    }
                    else
                    {
                        comment = string.Format(@"//TE_OPERATION {0} Pin = {1}",
                            log.Operation.Replace("_", ""), log.Interface);
                    }
                }
                else
                {
                    Match loopbackinfo = Regex.Match(comment, _reg_loopback, RegexOptions.IgnoreCase);
                    if (loopbackinfo.Success)
                    {
                        comment = ApplyLoopbackComment(comment, loopbackinfo, pinMapPins, ref loopBackFlag);
                    }

                    Match measinfo = Regex.Match(comment, _reg_meas, RegexOptions.IgnoreCase);
                    if (measinfo.Success)
                    {
                        comment = ApplyMeasPinsToComment(comment, measinfo, pinMapPins);
                    }
                }
                comment = AppendMeasNameComment(comment, log, benchLog, trimLogs);

                // ForceV MeasN case // TE_Operation  // ForceV Pin=VPP_OTP // ForceV = 5.3 // MeasN Pin = FakePin
                bool hasForceVPin = Regex.IsMatch(comment, @"ForceV\s+Pin\s*=", RegexOptions.IgnoreCase);
                bool hasForceV = Regex.IsMatch(comment, @"ForceV\s*=", RegexOptions.IgnoreCase);
                bool hasMeasN = Regex.IsMatch(comment, @"MeasN\s+Pin\s*=", RegexOptions.IgnoreCase);
                if (hasForceVPin && hasForceV && !hasMeasN)
                {
                    comment += " // MeasN Pin = FakePin";
                }

                if (isTrimStart)
                {
                    string regTrimName = @"MeasName\s*=\s*(?<name>\w+)";
                    trimMeasName = Regex.Match(comment, regTrimName, RegexOptions.IgnoreCase).Groups["name"].Value;
                }
                PrintCallSubroutine(patrows, infoV, "Global_TEOPEARTION_meas", lastVecFromPreBody, trimInfo, comment);
                isMeasCall = true;
                measCallExisted = true;
            }
            else if (Regex.IsMatch(log.Operation, "trim_start", RegexOptions.IgnoreCase))
            {
                SetTrimStartInfo(patrows, log, benchLog, ref trimInfo, ref readCapTrim);
            }

            return false;
        }

        private static string ApplyLoopbackComment(string comment, Match loopbackinfo, HashSet<string> pinMapPins,
            ref bool loopBackFlag)
        {
            loopBackFlag = true;
            comment = Regex.Replace(comment, @"(//LoopBack\s*)Path\s*=", "$1Pin=", RegexOptions.IgnoreCase);

            string type = loopbackinfo.Groups["type"].Value.ToLower();
            bool hasAtten = loopbackinfo.Groups["expectedattenuation"].Success;

            if ((type == "pin" || type == "path") && !hasAtten)
            {
                throw new Exception("Loopback pin/path must have expected_attenuation");
            }

            string[] parts = loopbackinfo.Groups["loopbackpath"].Value.Replace(" ", "").Split('-');

            if (parts.Length != 2)
            {
                throw new Exception("Loopback path format error");
            }

            string txPin = parts[0];
            string rxPin = parts[1];

            string tX_InsType = pinMapPins.FirstOrDefault(pin => pin.StartsWith(txPin + "_"));
            string rX_InsType = pinMapPins.FirstOrDefault(pin => pin.StartsWith(rxPin + "_"));

            string rlevl = hasAtten ? loopbackinfo.Groups["expectedattenuation"].Value.Replace("dB", "").Trim() : null;

            if (tX_InsType != null)
            {
                comment = comment.Replace(txPin + "-", tX_InsType + "-");
            }
            if (rX_InsType != null)
            {
                comment = comment.Replace("-" + rxPin, "-" + rX_InsType);
            }
            if ((type == "pin" || type == "path") && !comment.Contains("//TestType = LB"))
            {
                comment += "//TestType = LB";
            }
            else if (type == "disconnectpins" && !comment.Contains("//TestType = LB_Dis"))
            {
                comment += "//TestType = LB_Dis";
            }

            return comment;
        }

        private static string ApplyMeasPinsToComment(string comment, Match measinfo, HashSet<string> pinMapPins)
        {
            string pinname = measinfo.Groups["PinName"].Value;
            string pintype = measinfo.Groups["InstType"].Value;

            if (pinname.Contains("::"))
            {
                string[] allpins = pinname.Split(new string[] { "::" }, StringSplitOptions.None);
                foreach (string pin in allpins)
                {
                    if (pinMapPins.Contains(pin + "_SRC"))
                    {
                        comment = comment.Replace(pin, pin + "_SRC");
                    }
                }
            }
            else if (pinname.Contains("-"))
            {
                string[] str_lbp = pinname.Replace(" ", "").Split('-');
                string tX_InsType = pinMapPins.FirstOrDefault(pin => pin.StartsWith(str_lbp.First() + "_"));
                string rX_InsType = pinMapPins.FirstOrDefault(pin => pin.StartsWith(str_lbp.Last() + "_"));

                if (tX_InsType != null)
                {
                    comment = comment.Replace(str_lbp.First() + "-", tX_InsType + "-");
                }

                if (rX_InsType != null)
                {
                    comment = comment.Replace("-" + str_lbp.Last(), "-" + rX_InsType);
                }
            }
            else if (pinMapPins.Contains(pinname + "_LX"))
            {
                comment = comment.Replace(pinname, pinname + "_LX");
            }
            else if (pinMapPins.Contains(pinname + "_UW"))
            {
                comment = comment.Replace(pinname, pinname + "_UW");
            }

            return comment;
        }

        private static string AppendMeasNameComment(string comment, BenchLogItem log, BenchLogFile benchLog,
            List<BenchLogItem> trimLogs)
        {
            if (!Regex.IsMatch(comment, "MeasName", RegexOptions.IgnoreCase))
            {
                if (!string.IsNullOrEmpty(log.TestName))
                {
                    comment = ExpandTestNameByPins(comment, log.TestName);

                    if (!string.IsNullOrEmpty(log.LoLimit))
                    {
                        log.LoLimit = AppendUnitToLimit(log.LoLimit, log.Units);
                        comment = comment + string.Format("// LLimit = {0}", log.LoLimit);
                    }
                    if (!string.IsNullOrEmpty(log.HighLimit))
                    {
                        log.HighLimit = AppendUnitToLimit(log.HighLimit, log.Units);
                        comment = comment + string.Format("// HLimit = {0}", log.HighLimit);
                    }
                }
                else if (trimLogs.Count != 0)
                {
                    string regMeasInfo = @"(?<Type>\w+)\s+Pin\s*=\s*(?<PinName>[\w,:]+)";
                    string type =
                        Regex.Match(comment, regMeasInfo, RegexOptions.IgnoreCase).Groups["Type"].Value;
                    string pin =
                        Regex.Match(comment, regMeasInfo, RegexOptions.IgnoreCase).Groups["PinName"]
                            .Value.Replace("_", "");
                    string measName = string.Format("//MeasName = {0}_{1}_{2}_{3}_X_X_X_X_NV_VXXX",
                        benchLog.Type, benchLog.PatternSubName.Replace("_", ""), pin, type);
                    comment = comment + measName;
                }
            }

            return comment;
        }

        private static void SetTrimStartInfo(List<string> patrows, BenchLogItem log, BenchLogFile benchLog,
            ref string trimInfo, ref CapTrimInfo readCapTrim)
        {
            if (!benchLog.IsReadCapTrim)
            {
                string comment = string.Format("// TrimTarget = {0} // TrimType = {1} // BestCodeCalcFunc = {2}", log.Interface,
                    log.Address.Replace("0x", ""), log.RegData);
                if (!string.IsNullOrEmpty(log.LoLimit))
                {
                    comment = comment + string.Format("// LLimit = {0}{1}", log.LoLimit, log.Units);
                }

                if (!string.IsNullOrEmpty(log.HighLimit))
                {
                    comment = comment + string.Format("// HLimit = {0}{1}", log.HighLimit, log.Units);
                }

                trimInfo = comment;
            }
            else
            {
                patrows.Add(string.Format("//  TE_SETUP MiscInfo = BestCodeCalcFunc:{0};", log.RegData));
                readCapTrim = new CapTrimInfo { TrimTarget = log.Interface, BestCodeFunction = log.RegData };
                readCapTrim.LowLimit = string.Format("{0}{1}", log.LoLimit, log.Units);
                readCapTrim.HighLimit = string.Format("{0}{1}", log.HighLimit, log.Units);
            }
        }

        private void AppendMiscInfoRows(List<string> patrows, string trimMeasName, string block, string cppsetup,
            MappingItem patternName, bool isFW, bool readCapTrimFlag, bool loopBackFlag)
        {
            if (!string.IsNullOrEmpty(trimMeasName))
            {
                patrows.Add($"//TE_SETUP MiscInfo = TrimMeasName:{trimMeasName};");
            }
            if (PostProcessRows.Count > 0)
            {
                string miscComment = "//TE_SETUP MiscInfo = ";
                string funcname = GenType == GenerateType.HTOL ? VbtFunctionLibShared.RfHtolFunc : VbtFunctionLibShared.RfFunc;
                if (isFW)
                {
                    miscComment += string.Format("Func:{0};IsFW;", funcname);
                }
                else if (GenType == GenerateType.HTOL)
                {
                    miscComment += string.Format("Func:{0};", funcname);
                }
                else if (readCapTrimFlag)
                {
                    miscComment += string.Format("Func:{0};IsReadCapTrim;", funcname);
                }
                string pattern = patternName.Pattern;
                PostProcessRows.ForEach(p => p.BlockName = block);
                PostProcessRows.ForEach(p => p.SetupName = cppsetup);
                PostProcessRows.Where(p => string.IsNullOrEmpty(p.PatternName)).ToList().ForEach(p => p.PatternName = pattern);
                PostProcessRows.Select(p => p.BitWidth);
                int bits = PostProcessRows.Sum(p => int.Parse(p.BitWidth));
                miscComment = miscComment + string.Format("Post_Process:{0};DigCapBits:{1}", cppsetup, bits.ToString());
                patrows.Add(miscComment);
            }
            else if (loopBackFlag)
            {
                string miscComment = "//TE_SETUP MiscInfo = ";
                string funcname = GenType == GenerateType.HTOL ? VbtFunctionLibShared.RfHtolFunc : VbtFunctionLibShared.RfFunc;
                miscComment = miscComment + $"Func:{funcname};";
                patrows.Add(miscComment);
            }
        }

        private void AppendMatchLoopFiles(PatternResult result)
        {
            int matchindex = 0;
            foreach (KeyValuePair<string, List<string>> matchSet in MatchLoopInfo)
            {
                List<string> patrows = new List<string>();
                string srmName = string.Format("{0}_MatchSet", matchSet.Key + "_srm");
                CreateMatchLoop(patrows, srmName, matchSet.Key, matchindex, matchSet.Value);

                var patFile = new PatternFile
                {
                    FileName = matchSet.Key + ".atp"
                };

                patFile.Content.AddRange(patrows);
                result.Files.Add(patFile);

                if (!_itemInfo.PatSubr.Any(kv =>
                    kv.Key == matchSet.Key &&
                    kv.Value == srmName))
                {
                    _itemInfo.PatSubr.Add(
                        new KeyValuePair<string, string>(matchSet.Key, srmName)
                    );
                }
                matchindex++;
            }
        }

        public static string GenerateRelatedSCGH(string project, Dictionary<string, string> scghDic, Dictionary<string, List<string>> initSets, string outPath, string module)
        {
            if (string.IsNullOrEmpty(_project))
            {
                _project = project;
            }

            string resultPath = Path.Combine(outPath, string.Format("{0}_scgh_Template_{1}.xlsx", _project, DateTime.Now.ToString("yyyyMdd")));
            if (File.Exists(resultPath))
            {
                File.Delete(resultPath);
            }

            using (var ep = new ExcelPackage(new FileInfo(resultPath)))
            {
                //Block	Mode	Item	Application	INIT1	INIT2	INIT3	INIT4	INIT5	INIT6	PAYLOAD	USAGE (values are 1 or 0)

                ExcelWorksheet sheet = ep.Workbook.Worksheets.Add("SOC_HARD_IP_PROD_CHAR");
                sheet.Cells[1, 1].Value = "Block";
                sheet.Cells[1, 2].Value = "Mode";
                sheet.Cells[1, 3].Value = "Item";
                sheet.Cells[1, 4].Value = "Application";
                sheet.Cells[1, 5].Value = "INIT1";
                sheet.Cells[1, 6].Value = "INIT2";
                sheet.Cells[1, 7].Value = "INIT3";
                sheet.Cells[1, 8].Value = "INIT4";
                sheet.Cells[1, 9].Value = "INIT5";
                sheet.Cells[1, 10].Value = "PAYLOAD";
                sheet.Cells[1, 11].Value = "USAGE (values are 1 or 0)";
                sheet.Cells[1, 12].Value = "Comment";
                sheet.Cells[1, 13].Value = "";
                int rowindex = 2;

                var referenceInits = initSets.SelectMany(p => p.Value).Distinct().ToList();
                foreach (KeyValuePair<string, string> scghinfo in scghDic)
                {
                    if (referenceInits.Any(p => p.Equals(scghinfo.Value, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    List<string> itemSgmt = scghinfo.Key.Split('_').ToList();
                    string itemName = "";
                    if (Regex.IsMatch(itemSgmt[0], @"^\d+$"))
                    {
                        itemName = string.Join("_", itemSgmt.GetRange(1, itemSgmt.Count - 1));
                    }
                    else
                    {
                        itemName = scghinfo.Key;
                    }

                    sheet.Cells[rowindex, 1].Value = module;
                    sheet.Cells[rowindex, 2].Value = module;
                    sheet.Cells[rowindex, 3].Value = itemName;
                    sheet.Cells[rowindex, 4].Value = "Production";
                    List<string> initset = initSets.ContainsKey(scghinfo.Value) ? initSets[scghinfo.Value] : new List<string>();
                    int init_count = initset.Count > 5 ? 5 : initset.Count;
                    #region set init pattern
                    for (int i = 0; i < init_count; i++)
                    {
                        sheet.Cells[rowindex, 5 + i].Value = initset[i];
                    }
                    #endregion
                    sheet.Cells[rowindex, 10].Value = scghinfo.Value;
                    sheet.Cells[rowindex, 11].Value = "1";
                    rowindex++;
                }
                ep.Save();
                ep.Dispose();
            }

            return resultPath;
        }

        public static void GeneratePatSetInfo(string logDirect, Dictionary<string, string> patsetallDic, Dictionary<string, List<string>> patsetsubrDic, string specifyPatternFolder)
        {
            string patsetAllName = Path.Combine(logDirect, "PatSets_All.txt");
            var patsetAll = new PatSetSheet("PatSets_All.txt");
            if (File.Exists(patsetAllName))
            {
                File.Delete(patsetAllName);
            }
            foreach (KeyValuePair<string, string> patsetall in patsetallDic)
            {
                var patset = new PatSet();
                patset.PatSetName = patsetall.Key;
                var row = new PatSetRow();
                row.File = string.Format(@".\PATTERN\{1}\{0}.PAT:{0}", patsetall.Value, specifyPatternFolder);
                row.PatternSet = patsetall.Key;
                row.Burst = "no";

                patset.AddRow(row);
                patsetAll.Rows.Add(patset);
            }
            patsetAll.Write(patsetAllName, "2.3");
            string patsetSubrName = Path.Combine(logDirect, "Pattern_Subroutine.txt");
            if (File.Exists(patsetSubrName))
            {
                File.Delete(patsetSubrName);
            }

            var swPatsetSubr = new StreamWriter(patsetSubrName);
            swPatsetSubr.WriteLine("DTPatternSubroutineSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1	Pattern Subroutine");
            swPatsetSubr.WriteLine();
            swPatsetSubr.WriteLine("	Pattern Filename	Comment");
            foreach (KeyValuePair<string, List<string>> patsetsub in patsetsubrDic)
            {
                foreach (string subrItem in patsetsub.Value)
                {
                    swPatsetSubr.WriteLine(@"	.\PATTERN\{2}\{0}.PAT:{1}", patsetsub.Key, subrItem, specifyPatternFolder);
                }
            }
            swPatsetSubr.Close();
        }

        public static void GenerateLUT(string logDirect, List<LutItem> luts)
        {
            if (luts.Count == 0)
            {
                return;
            }

            string lutName = Path.Combine(logDirect, "RFLookUpTable.txt");
            if (File.Exists(lutName))
            {
                File.Delete(lutName);
            }

            var swAll = new StreamWriter(lutName);
            foreach (LutItem lut in luts)
            {
                foreach (string line in lut.PrintLut())
                {
                    swAll.WriteLine(line);
                }
            }

            swAll.Close();
        }

        public static void GenerateCPPTable(string logDirect, List<PostProcessSheetRow> postitems)
        {
            if (postitems.Count > 0)
            {
                var writer = new PostProcessSheet();
                writer.RowList = postitems;
                writer.Write(logDirect);
            }
        }

        private void UpdateTableValue(string path, string refaddress, Dictionary<string, Register> totalRegisters,
            string payload, string projName, string siliconVer, bool isFullSweep, bool isDebugMode = false)
        {
            string refLog = refaddress.Split('#')[0];
            if (string.IsNullOrEmpty(refLog))
            {
                return;
            }

            foreach (string logFile in Directory.GetFiles(path, "*.csv"))
            {
                bool isFindTarget = true;
                if (!Path.GetFileNameWithoutExtension(logFile).Equals(refaddress))
                {
                    continue;
                }

                MappingItem initname = GetNamingInfo(logFile, GenType, projName, siliconVer, isFullSweep, isDebugMode);
                string init = initname.Pattern;
                _itemInfo.InitDictionary.Add(init);
                List<BenchLogFile> loginfos = ReadLogFile(logFile, isFullSweep);
                List<BenchLogItem> logs = loginfos.First().Logs;

                _ = new Dictionary<string, string>();
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
                            UpdateTableValue(path, log.Interface, totalRegisters, payload, projName, siliconVer, isFullSweep, isDebugMode);
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

        }
        private static string GetDigSrcRegisterName(CaptureInfo log)
        {
            return $"Addr{log.Address}_M{log.MSB}_L{log.LSB}";
        }

        private static string GetDigSrcRegisterName(BenchLogItem log)
        {
            return $"Addr{log.Address}_M{log.MSB}_L{log.LSB}";
        }

        private Register UpdateRegisterValue(BenchLogItem log, Dictionary<string, Register> registers, string relatedVar, bool isReg)
        {
            string data = log.FieldVal;
            if (string.IsNullOrEmpty(data))
            {
                data = "0x0";
            }

            string register = log.RegFieldName.Split('.').Last();
            if (string.IsNullOrEmpty(register))
            {
                register = $"Addr{log.Address}_M{log.MSB}_L{log.LSB}";
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
                ApplyFieldValToRegister(log, result, data, register, isReg, isNeedShift, shiftbit);
            }
            else
            {
                var overWriteList = new Dictionary<string, string>();
                char[] sourceStr = ApplyRegDataToRegister(log, result, isNeedShift, shiftbit, overWriteList);
                LogOverWrittenFields(log, overWriteList);
                result.Data = new string(sourceStr);
            }

            return result;
        }

        private void ApplyFieldValToRegister(BenchLogItem log, Register result, string data, string register, bool isReg,
            bool isNeedShift, int shiftbit)
        {
            string binStr = ConvertValToBin(data, int.Parse(result.Length), isNeedShift);
            if (isReg)
            {
                char[] binArr = binStr.ToCharArray();
                for (int i = 0; i < PatBit; i++)
                {
                    if (result.Data.Substring(i, 1) == "D")
                    {
                        binArr[i] = 'D';
                    }
                }
                result.Data = new string(binArr);
            }
            else
            {
                string regField = @"(?<name>\w+)(?<partial>\[.*\])*";

                _ = Regex.Match(register, regField, RegexOptions.IgnoreCase).Groups["name"].Value;

                _ = Regex.Match(register, regField, RegexOptions.IgnoreCase).Groups["partial"].Value;
                char[] regdataArr = result.Data.ToCharArray();
                char[] binArr = binStr.ToCharArray();
                int index = 0;
                for (int i = PatBit - 1 + shiftbit - log.LSB; i >= PatBit - 1 + shiftbit - log.MSB; i--)
                {
                    if (!string.IsNullOrEmpty(log.FuseName))
                    {
                        regdataArr[i] = 'D';
                    }
                    else
                    {
                        regdataArr[i] = binArr[binArr.Length - 1 - index];
                    }

                    index++;
                }
                result.Data = new string(regdataArr);
            }
        }

        private char[] ApplyRegDataToRegister(BenchLogItem log, Register result, bool isNeedShift, int shiftbit,
            Dictionary<string, string> overWriteList)
        {
            char[] sourceStr = result.Data.ToCharArray();
            char[] regDataStr = ConvertValToBin(log.RegData, int.Parse(result.Length), isNeedShift).ToCharArray();
            var regSrc = new Regex(@"_M(?<msb>\d+)_L(?<lsb>\d+)", RegexOptions.Compiled);

            for (int i = 0; i < PatBit; i++)
            {
                if (!IsAllowOverWrite)
                {
                    if (sourceStr[i] == 'D')
                    {
                        ;
                    }
                    else if (PatBit - 1 - i - shiftbit <= log.MSB && PatBit - 1 - i - shiftbit >= log.LSB)
                    {
                        if (!string.IsNullOrEmpty(log.FuseName))
                        {
                            sourceStr[i] = 'D';
                        }
                        else
                        {
                            sourceStr[i] = regDataStr[i];
                        }
                    }
                    else
                    {
                        sourceStr[i] = regDataStr[i];
                    }
                }
                else
                {
                    if (PatBit - 1 - i - shiftbit <= log.MSB && PatBit - 1 - i - shiftbit >= log.LSB)
                    {
                        if (sourceStr[i] == 'D')
                        {
                            CollectOverWrittenFields(log, result, overWriteList, regSrc, i, shiftbit);
                        }
                        if (!string.IsNullOrEmpty(log.FuseName))
                        {
                            sourceStr[i] = 'D';
                        }
                        else
                        {
                            sourceStr[i] = regDataStr[i];
                        }
                    }
                    else if (sourceStr[i] == 'D')
                    {
                        ;
                    }
                    else
                    {
                        sourceStr[i] = regDataStr[i];
                    }
                }
            }

            return sourceStr;
        }

        private void CollectOverWrittenFields(BenchLogItem log, Register result, Dictionary<string, string> overWriteList,
            Regex regSrc, int i, int shiftbit)
        {
            foreach (KeyValuePair<string, List<string>> srcinfo in result.SrcDic)
            {
                if (overWriteList.ContainsKey(srcinfo.Key))
                {
                    continue;
                }

                int msb = int.Parse(regSrc.Match(srcinfo.Key).Groups["msb"].Value);
                int lsb = int.Parse(regSrc.Match(srcinfo.Key).Groups["lsb"].Value);
                if (PatBit - 1 - i - shiftbit <= msb && 31 - i - shiftbit >= lsb)
                {
                    if (!log.FuseName.Equals(srcinfo.Value.Last(), StringComparison.OrdinalIgnoreCase))
                    {
                        overWriteList.Add(srcinfo.Key, srcinfo.Value.Last());
                    }
                }
            }
        }

        private void LogOverWrittenFields(BenchLogItem log, Dictionary<string, string> overWriteList)
        {
            if (overWriteList.Count > 0)
            {
                foreach (KeyValuePair<string, string> srcinfo in overWriteList)
                {
                    if (!string.IsNullOrEmpty(log.FuseName))
                    {
                        if (log.RawData.ToLower().IndexOf("write_trim") >= 0)
                        {
                            continue;
                        }
                        else
                        {
                            _swlog.WriteLine("Field: \"{0}\" with Fuse: {1} has been overwritten by {2}.", srcinfo.Key, srcinfo.Value, log.FuseName);
                        }
                    }
                    else if (!string.IsNullOrEmpty(log.FieldVal))
                    {
                        _swlog.WriteLine("Field: \"{0}\" with Fuse: {1} has been overwritten by value:{2}.", srcinfo.Key, srcinfo.Value, log.FieldVal);
                    }
                    else
                    {
                        _swlog.WriteLine("Field: \"{0}\" with Fuse: {1} has been overwritten by regData:{2}.", srcinfo.Key, srcinfo.Value, log.RegData);
                    }
                }
            }
        }

        private void UpdateSourceInfo(BenchLogItem log, Register register)
        {
            if (register.SrcDic.Count == 0 || !IsAllowOverWrite)
            {
                return;
            }

            var regMSBLSB = new Regex(@"_M(?<msb>\d+)_L(?<lsb>\d+)", RegexOptions.Compiled);
            string targetPos = string.Format("_M{0}_L{1}", log.MSB, log.LSB);
            string targetReg = register.SrcDic.Keys.FirstOrDefault(p => Regex.IsMatch(p, targetPos, RegexOptions.IgnoreCase));
            if (targetReg != null)
            {
                return;
            }

            var result = new Dictionary<string, List<string>>();
            foreach (KeyValuePair<string, List<string>> item in register.SrcDic)
            {
                int msb = int.Parse(regMSBLSB.Match(item.Key).Groups["msb"].Value);
                int lsb = int.Parse(regMSBLSB.Match(item.Key).Groups["lsb"].Value);
                if (lsb > log.MSB || msb < log.LSB)
                {
                    result.Add(item.Key, item.Value);
                }
            }
            register.SrcDic = result;
        }

        private Register SetReadCapValue(BenchLogItem log, Dictionary<string, Register> registers,
            string relatedVar)
        {
            _ = log.RegFieldName.Split('.').Last();

            _ = log.LSB;

            _ = log.MSB;

            if (!string.IsNullOrEmpty(relatedVar) && Regex.IsMatch(relatedVar, "^0x", RegexOptions.IgnoreCase))
            {
            }

            Register result = registers.ContainsKey(log.Address) ? registers[log.Address] : null;

            if (result == null)
            {
                int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
                bool isNeedShift = PatBit == 64 && addrQuot % 2 != 0;
                result = CreateRegister(log, "0x0", isNeedShift);
                registers.Add(log.Address, result);
            }

            return result;
        }

        private Register CreateRegister(BenchLogItem log, string defValue, bool isNeedShift)
        {
            var result = new Register();
            string binStr;
            if (string.IsNullOrEmpty(defValue) || !defValue.StartsWith("0x"))
            {
                if (string.IsNullOrEmpty(defValue) && string.IsNullOrEmpty(log.RegData))
                {
                    binStr = ConvertValToBin("0x0", PatBit, isNeedShift);
                }
                else
                {
                    binStr = ConvertValToBin(log.RegData, PatBit, isNeedShift);
                }
            }
            else
            {
                binStr = ConvertValToBin(defValue, PatBit, isNeedShift);
            }

            result.Address = log.Address;
            result.Length = PatBit.ToString();
            result.Data = binStr;
            return result;
        }

        private Register SetReadValue(BenchLogItem log, Dictionary<string, Register> registers)
        {
            string address = log.Address;

            _ = log.RegFieldName.Split('.').Last();

            _ = log.FieldVal;

            _ = log.LSB;

            _ = log.MSB;
            Register result = registers.ContainsKey(address) ? registers[address] : null;

            if (result == null)
            {
                int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
                bool isNeedShift = PatBit == 64 && addrQuot % 2 != 0;
                result = CreateRegister(log, "0x0", isNeedShift);
                registers.Add(log.Address, result);
            }

            return result;
        }

        public static string ExpandTestNameByPins(string comment, string logTestName)
        {
            Match match = Regex.Match(comment, @"(?:Pin|path|DisconnectPins)\s*=\s*([^/]+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                throw new Exception("Pin not found: " + comment);
            }

            string measPins = match.Groups[1].Value.Trim();
            List<string> listmeasPins = measPins.Split(',').Select(p => p.Trim()).ToList();

            listmeasPins = listmeasPins.Select(p => p.Replace("_", "")).ToList();
            List<string> listlogTestName = logTestName.Split(',').ToList();
            string[] templogTestName = logTestName.Split('_');
            string measName = string.Format("//MeasName = {0}", logTestName);
            if (listmeasPins.Count != listlogTestName.Count)
            {
                listlogTestName.Clear();
                List<string> measlogTestName = new List<string>();
                for (int i = 0; i < listmeasPins.Count; i++)
                {
                    templogTestName[2] = listmeasPins[i];
                    logTestName = string.Join("_", templogTestName);
                    measlogTestName.Add(logTestName);
                }
                measName = string.Format("//MeasName = {0}", string.Join(",", measlogTestName));
            }
            return comment + measName;
        }

        public static string AppendUnitToLimit(string input, string unit)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }
            return string.Join(",", input.Split(',').Select(x => x + unit));
        }
        private static string ConvertValToBin(string data, int size, bool isShiftToHight = false)
        {
            try
            {
                string result = Regex.Match(data, @"(0x)*(?<data>\w+)", RegexOptions.IgnoreCase).Groups["data"].Value;

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
                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void CreateSubRoutine(List<string> patrows, string srmName, string subrName)
        {
            _subrP.Write(patrows, srmName, subrName);
        }

        private static void CreateMatchLoop(List<string> patrows, string srmName, string subrName, int index, List<string> datas)
        {
            MatchLoopPattern.Write(patrows, srmName, subrName, index, datas);
        }

        private void PrintswBody(TextWriter sw, VectorInfo info, string patternname, bool isSource, bool isCap, bool isMeasCall, HashSet<string> matchLoopSAddressSRM)
        {
            _bodyP.WriteMerge(sw, patternname, info, isSource, isCap, isMeasCall, matchLoopSAddressSRM);
        }

        private string GetLastVec()
        {
            return _bodyP.RawDataList.FindLast(raw => raw.Contains(">"));
        }

        private static void PrintCallSubroutine(List<string> patrows, VectorInfo info, string name, string lastVec, string trimInfo, string comment = "")
        {
            string regPin = @"Pin\s*=\s*(?<pName>[\w\:\;\-]+)";
            /* > tsetJTAG 1 0 0 1 1 1 1 1 X 0 0 0 X X X XXXX00 X XXXXXXX X XXXX XXXX X X X X X X XXXXXXXX XXXXXXX X XXXXXXX X X XXXX X XXX X XXXX X XXX X X X X X X X XXXX XXXX X XX XXXX XXX XXXX X XXXX X XXXXX X X XXXXX X X X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX X XX X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XX XX XXXX 00 ; // 5272671 V:59232 C:126544 
call PP_BTCA0_S_PLLP_AN_RFD4_PFF_JTG_UNS_ALLFRV_SI_WL5GTRX_PLLPNST_digcapsrc  > tsetJTAG 1 0 0 1 1 1 1 1 X 0 0 0 X X X XXXX00 X XXXXXXX X XXXX XXXX X X X X X X XXXXXXXX XXXXXXX X XXXXXXX X X XXXX X XXX X XXXX X XXX X X X X X X X XXXX XXXX X XX XXXX XXX XXXX X XXXX X XXXXX X X XXXXX X X X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX X XX X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XX XX XXXX 00 ; // 5272713 V:59233 C:126545 //TE_OPERATION MeasV Pin = ARF_PADIO_TRXPLL_5G_VCO_TP //MeasStoreName = 5gtrx_pll_vtune 
            */

            string pin = Regex.Match(comment, regPin, RegexOptions.IgnoreCase).Groups["pName"].Value;

            string pReplacePin = pin.Replace(";", ",");
            if (!string.IsNullOrEmpty(pin))
            {
                comment = comment.Replace(pin, pReplacePin);
            }

            if (!string.IsNullOrEmpty(trimInfo))
            {
                comment = comment + " " + trimInfo;
            }

            string patrow;
            if (!string.IsNullOrEmpty(info.LastVector))
            {
                patrow = info.LastVector;
                if (IsAddComment)
                {
                    patrow += info.GetVectorInfo();
                    info.Update();
                }
                patrows.Add(patrow);

                patrow = string.Format("call {0} {1}// {2}", name, info.LastVector, comment);
                if (IsAddComment)
                {
                    patrow += info.GetVectorInfo();
                    info.Update();
                }
                patrows.Add(patrow);
            }
            else
            {

                patrow = lastVec;
                if (IsAddComment)
                {
                    patrow += info.GetVectorInfo();
                    info.Update();
                }
                patrows.Add(patrow);

                patrow = string.Format("call {0} {1}// {2}", name, lastVec, comment);
                if (IsAddComment)
                {
                    patrow += info.GetVectorInfo();
                    info.Update();
                }
                patrows.Add(patrow);
            }
        }

        private void PrintWait(List<string> patrows, VectorInfo info, string logintfac, string lastVec)
        {
            /* > tsetJTAG 1 0 0 1 1 1 1 1 X 0 0 0 X X X XXXX00 X XXXXXXX X XXXX XXXX X X X X X X XXXXXXXX XXXXXXX X XXXXXXX X X XXXX X XXX X XXXX X XXX X X X X X X X XXXX XXXX X XX XXXX XXX XXXX X XXXX X XXXXX X X XXXXX X X X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX X XX X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XX XX XXXX 00 ; // 5272671 V:59232 C:126544 
call PP_BTCA0_S_PLLP_AN_RFD4_PFF_JTG_UNS_ALLFRV_SI_WL5GTRX_PLLPNST_digcapsrc  > tsetJTAG 1 0 0 1 1 1 1 1 X 0 0 0 X X X XXXX00 X XXXXXXX X XXXX XXXX X X X X X X XXXXXXXX XXXXXXX X XXXXXXX X X XXXX X XXX X XXXX X XXX X X X X X X X XXXX XXXX X XX XXXX XXX XXXX X XXXX X XXXXX X X XXXXX X X X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX X XX X XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XX XX XXXX 00 ; // 5272713 V:59233 C:126545 //TE_OPERATION MeasV Pin = ARF_PADIO_TRXPLL_5G_VCO_TP //MeasStoreName = 5gtrx_pll_vtune 
             */

            string vecinfo = string.IsNullOrEmpty(info.LastVector) ? lastVec : info.LastVector;

            double defV = Math.Ceiling(_freq * 1e6 * double.Parse(logintfac));

            string patrow;
            if (defV > 65536)
            {
                int tmpQ;
                for (int j = 0; j < Math.DivRem((int)defV, 65536, out tmpQ); j++)
                {
                    patrow = string.Format("repeat 65536 {0}", vecinfo);
                    if (IsAddComment)
                    {
                        patrow += info.GetVectorInfo();
                        info.Update();
                    }
                    patrows.Add(patrow);
                }
                defV = tmpQ;

            }

            patrow = string.Format("repeat {0} {1}", defV, vecinfo);
            if (IsAddComment)
            {
                patrow += info.GetVectorInfo();
                info.Update();
            }
            patrows.Add(patrow);
        }

        //WRITE_SRC WRITE
        private void PrintWrite(List<string> patrows, BenchLogItem log, Dictionary<string, Register> registers,
            VectorInfo info, bool isFuse, ref bool isSource, bool isReg = false)
        {
            Register register = null;
            var digsrcList = new List<string>();

            if (log.RegData.Length > 10 && PatBit == 32)
            {
                throw new Exception($"Error occurs using 64 bits REG_DATA : {log.RegData} in 32bits pattern");
            }

            foreach (BenchLogItem src in log.SrcInfo)
            {
                if (!string.IsNullOrEmpty(src.FuseName))
                {
                    register = UpdateRegisterValue(src, registers, src.FuseName, isReg);
                }
                else
                {
                    register = UpdateRegisterValue(src, registers, src.Default, isReg);
                }

                UpdateSourceInfo(log, register);
                if (!string.IsNullOrEmpty(src.FuseName) && isFuse)
                {
                    PrintTrimFuseSource(patrows, log, src, register, digsrcList, ref isSource);
                }
                else if (GenType != GenerateType.HTOL && !string.IsNullOrEmpty(src.FuseName) &&
                   !Regex.IsMatch(src.FuseName, "^0x", RegexOptions.IgnoreCase))
                {
                    AddDigSrcForFuseName(log, src, register, digsrcList);
                }
                else if (register != null && string.IsNullOrEmpty(src.FuseName))
                {
                    AddDigSrcForUsedSource(log, register, digsrcList);
                }
            }
            if (digsrcList.Count > 0 && register.Data.Contains("D"))
            {
                isSource = true;
                foreach (string digsrc in digsrcList)
                {
                    patrows.Add($"// TE_SETUP DigSrcAssignment = {digsrc}");
                }
            }
            string comInfo = $"{log.Operation}#{log.Address}_L{log.LSB}_M{log.MSB}#{log.RegData}#{log.FieldVal}#";
            patrows.Add("//CMD Start " + comInfo);
            string addInBin = "";
            if (uint.TryParse(log.Address.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out uint num))
            {
                addInBin = Convert.ToString(num, 2);
            }

            int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
            if (_fullAddrFor64InBin.Contains(addInBin) && addrQuot % 2 == 0)
            {
                _write64P.Write(patrows, log, register, info, _swlog);
            }
            else
            {
                _writeP.Write(patrows, log, register, info, _swlog);
            }

            patrows.Add("//CMD End " + comInfo);
        }

        private void PrintTrimFuseSource(List<string> patrows, BenchLogItem log, BenchLogItem src, Register register,
            List<string> digsrcList, ref bool isSource)
        {
            isSource = true;
            string srcReg = GetDigSrcRegisterName(src);
            patrows.Add($"// TE_SETUP TrimFuseName = {log.FuseName} //TrimRegName = {srcReg}");
            foreach (KeyValuePair<string, List<string>> reg in register.SrcDic)
            {
                if (!_usedDic.ContainsKey(reg.Key))
                {
                    _usedDic.Add(reg.Key, new List<string>());
                }

                if (!_usedDic[reg.Key].Contains(reg.Value.Last()))
                {
                    _usedDic[reg.Key].Add(reg.Value.Last());
                    digsrcList.Add($"{reg.Key}:{reg.Value.Last()}");
                }
            }
            if (!register.SrcDic.ContainsKey(srcReg))
            {
                register.SrcDic.Add(srcReg, new List<string>());
            }

            register.SrcDic[srcReg].Add(log.FuseName);
        }

        private void AddDigSrcForFuseName(BenchLogItem log, BenchLogItem src, Register register, List<string> digsrcList)
        {
            string regName = GetDigSrcRegisterName(src);
            string value = src.FuseName.Replace(":", ";");
            if (!register.SrcDic.ContainsKey(regName))
            {
                register.SrcDic.Add(regName, new List<string>());
            }

            register.SrcDic[regName].Add(value);

            foreach (KeyValuePair<string, List<string>> reg in register.SrcDic)
            {
                if (!_usedDic.ContainsKey(reg.Key))
                {
                    _usedDic.Add(reg.Key, new List<string>());
                }

                if (!_usedDic[reg.Key].Contains(reg.Value.Last()))
                {
                    _usedDic[reg.Key].Add(reg.Value.Last());
                    string source_reg = Register.GetAssignReg(reg.Key, reg.Value, log);
                    digsrcList.Add($"{source_reg}:{reg.Value.Last()}");
                }
            }
        }

        private void AddDigSrcForUsedSource(BenchLogItem log, Register register, List<string> digsrcList)
        {
            foreach (KeyValuePair<string, List<string>> reg in register.SrcDic)
            {
                string regMSBLSB = @"M(?<msb>\d+)_L(?<lsb>\d+)";
                string lsb = Regex.Match(reg.Key, regMSBLSB, RegexOptions.IgnoreCase).Groups["lsb"].Value;
                string msb = Regex.Match(reg.Key, regMSBLSB, RegexOptions.IgnoreCase).Groups["msb"].Value;
                if (log.MSB >= int.Parse(msb) && log.LSB <= int.Parse(lsb))
                {
                    continue;
                }

                if (!_usedDic.ContainsKey(reg.Key))
                {
                    _usedDic.Add(reg.Key, new List<string>());
                }

                if (!_usedDic[reg.Key].Contains(reg.Value.Last()))
                {
                    _usedDic[reg.Key].Add(reg.Value.Last());
                    string source_reg = Register.GetAssignReg(reg.Key, reg.Value, log);
                    digsrcList.Add($"{source_reg}:{reg.Value.Last()}");
                }
            }
        }

        private void PrintRead(List<string> patrows, BenchLogItem log, Dictionary<string, Register> registers,
            VectorInfo info, bool isTrimStart, PatternGenItem itemInfo)
        {

            if (log.RegData.Length > 10 && PatBit == 32)
            {
                throw new Exception($"Error occurs using 64 bits REG_DATA : {log.RegData} in 32bits pattern");
            }
            Register register = SetReadCapValue(log, registers, log.Default);
            foreach (CaptureInfo readinfo in log.Capinfos)
            {
                string reg = GetDigSrcRegisterName(readinfo);
                if (!register.CapDic.ContainsKey(reg))
                {
                    register.CapDic.Add(reg, "");
                }
            }
            var cpp = new List<PostProcessSheetRow>();
            string comInfo = $"{log.Operation}#{log.Address}_L{log.LSB}_M{log.MSB}#{log.RegData}#{log.FieldVal}#";
            patrows.Add("//CMD Start " + comInfo);
            string addInBin = "";
            if (uint.TryParse(log.Address.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out uint num))
            {
                addInBin = Convert.ToString(num, 2);
            }

            int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
            if (_fullAddrFor64InBin.Contains(addInBin) && addrQuot % 2 == 0)
            {
                _read64P.Write(patrows, log, register, info, ref cpp, _swlog, isTrimStart, itemInfo);
            }
            else
            {
                _readP.Write(patrows, log, register, info, ref cpp, _swlog, isTrimStart, itemInfo);
            }

            patrows.Add("//CMD End " + comInfo);


            foreach (CaptureInfo readinfo in log.Capinfos)
            {
                if (!string.IsNullOrEmpty(readinfo.FuseName))
                {
                    foreach (string post in readinfo.PostProcess.Split(';'))
                    {
                        ////TE_SETUP MiscInfo =
                        if (post.Split(':')[0].Equals("storename", StringComparison.OrdinalIgnoreCase))
                        {
                            patrows.Add($"//TE_SETUP MiscInfo = SetEfuse : {post.Split(':')[1]} = {readinfo.FuseName}");
                        }
                    }
                }
            }

            if (cpp.Count > 0)
            {
                PostProcessRows.AddRange(cpp);
            }
        }

        private List<string> PrintWaitFor(BenchLogItem log, Dictionary<string, Register> registers,
            VectorInfo info, Logger logger)
        {
            Register register = SetReadValue(log, registers);
            if (register == null)
            {
                logger.Info(string.Format("Address not found? {0}", log.Address));
            }

            return _matchP.Write(log, register, info, _swlog);

        }

        private static Dictionary<string, Register> FilterBanks(Dictionary<string, Register> banks, List<string> addressList)
        {
            var result = new Dictionary<string, Register>();
            foreach (string address in addressList)
            {
                Register tbank = banks.ContainsKey(address) ? banks[address] : null;

                if (tbank != null)
                {
                    tbank.Address = address;
                    result.Add(address, tbank);
                }
            }

            return result;
        }

        public static string GetStringName(string data)
        {
            return Regex.Match(data, "\"(?<str>.*)\"", RegexOptions.IgnoreCase).Groups["str"].Value;
        }

        public static List<string> GetRegisters(string type)
        {
            var result = new List<string>();
            string reg_bracketRemoved = @"\[(?<data>.*)\]";
            string reg_infos = Regex.Match(type.Replace("\"", ""), reg_bracketRemoved, RegexOptions.IgnoreCase).Groups["data"].Value;
            foreach (string reg_info in Regex.Split(reg_infos, @"[\[\]]", RegexOptions.IgnoreCase))
            {
                if (string.IsNullOrEmpty(reg_info.Trim(',')))
                {
                    continue;
                }

                int bitsize = 0;
                for (int i = 0; i < type.Split(',').Count(); i++)
                {
                    try
                    {
                        switch (i)
                        {
                            case 0:
                                bitsize = int.Parse(reg_info.Split(':').First().Split(',')[i].Trim());
                                break;
                            case 2:
                                string value = Convert.ToString(Convert.ToInt64(reg_info.Split(',')[i].Trim()), 2).PadLeft(bitsize, '0');
                                result.Add(value);
                                break;
                            default:
                                break;
                        }
                    }
                    catch (Exception)
                    {
                        ;
                    }
                }
            }
            return result;
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

        private List<BenchLogFile> ReadLogFile(string originpath, bool isFullSweep, bool isDebugMode = false, string outputDir = null)
        {
            var result = new List<BenchLogFile>();
            var loginfo = new BenchLogFile();
            string logDirect = Path.GetDirectoryName(outputDir);

            Dictionary<string, string> efusebigDefRowsToDic = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            efusebigDefRowsToDic = EfusebitDefRows
                            .Where(r => r.Count >= 2 && !string.IsNullOrWhiteSpace(r[0]))
                            .Select(r => new { Key = r[0].Trim(), Value = (r[3] ?? string.Empty).Trim() })
                            .GroupBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First().Value);
            loginfo.Offset = -1;
            string tmp = originpath;
            bool isNeedDelete = false;
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
            bool isTrim = false;
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

                    if (line.Contains("WRITE,JTAG,0x4108e7a0,31,0,,0x304000,,,,,,,,"))
                    {
                        ;
                    }
                    if (loginfo.Offset == -1 && line.ToUpper().Contains("OPERATION"))
                    {
                        loginfo.Offset =
                            line.Split(',')
                                .ToList()
                                .FindIndex(p => p.Equals("OPERATION", StringComparison.OrdinalIgnoreCase));
                        loginfo.GetHeader(line);

                    }
                    var log = new BenchLogItem(line, loginfo, row, isFullSweep, ref isFirst);
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
                        isTrim = false;
                        List<string> logsplit = log.RawData.Split(',').ToList();
                        loginfo = loginfo.Copy();
                        loginfo.PatternSubName = logsplit[3];
                        loginfo.Type = logsplit[2];
                        loginfo.Interface = logsplit[1];
                        loginfo.Version = string.IsNullOrEmpty(logsplit[4]) ? "1" : logsplit[4];
                        loginfo.DateCode = isDebugMode ? "DEBUGMODE" : string.IsNullOrEmpty(logsplit[5]) ? "" : logsplit[5];
                        result.Add(loginfo);

                        continue;
                    }
                    if (Regex.IsMatch(log.Operation, "OPERATION|print", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }
                    AddLogToBenchLog(ref log, line, loginfo, efusebigDefRowsToDic, row, isFullSweep, ref isFirst);

                    UpdateTrimState(log, loginfo, isFullSweep, ref isTrim);
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

        private void AddLogToBenchLog(ref BenchLogItem log, string line, BenchLogFile loginfo,
            Dictionary<string, string> efusebigDefRowsToDic, int row, bool isFullSweep, ref bool isFirst)
        {
            if (log.Operation.Equals("WAIT", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    loginfo.Logs.Add(log);
                }
                catch (Exception)
                {
                    ;
                }
            }
            else if (log.Operation.Equals("Read_cap", StringComparison.OrdinalIgnoreCase))
            {
                if (loginfo.Logs.Count > 0)
                {
                    BenchLogItem logCap = loginfo.Logs.Last(p => !p.Operation.Equals("calc", StringComparison.OrdinalIgnoreCase));
                    if (IsNeedMergeReadCap(log, logCap))
                    {
                        logCap.Capinfos.AddRange(log.Capinfos);
                        IsSwapReedCap(logCap);
                    }
                    else
                    {
                        loginfo.Logs.Add(log);
                    }
                }
                else
                {
                    loginfo.Logs.Add(log);
                }
            }
            else if (Regex.IsMatch(log.Operation, "write_src", RegexOptions.IgnoreCase))
            {
                int bitwidth = log.MSB - log.LSB + 1;
                if (!efusebigDefRowsToDic.ContainsKey(log.FuseName.ToUpper()) && !loginfo.WriteSrcRows.ContainsKey(log.FuseName.ToUpper()))
                {
                    loginfo.WriteSrcRows.Add(log.FuseName, bitwidth.ToString());
                }

                if (loginfo.Logs.Count > 0)
                {
                    BenchLogItem logSrc = loginfo.Logs.Last();
                    if (IsNeedMergeWriteSrc(log, logSrc))
                    {
                        logSrc.SrcInfo.Add(log);
                    }
                    else
                    {
                        loginfo.Logs.Add(log);
                    }
                }
                else
                {
                    loginfo.Logs.Add(log);
                }
            }
            else if (Regex.IsMatch(log.Operation, "wait_for", RegexOptions.IgnoreCase))
            {
                if (GenType == GenerateType.HTOL)
                {
                    string logstr = "WAIT,0.1,,,,,,,,,,,,,";
                    loginfo.Logs.Add(new BenchLogItem(logstr, loginfo, row, isFullSweep, ref isFirst));

                    logstr = Regex.Replace(line, "wait_for", "read_compare", RegexOptions.IgnoreCase);
                    loginfo.Logs.Add(new BenchLogItem(logstr, loginfo, row, isFullSweep, ref isFirst));
                }
                else
                {
                    loginfo.Logs.Add(log);
                    string newLine = Regex.Replace(line, "wait_for", "read_cap", RegexOptions.IgnoreCase);
                    log = new BenchLogItem(newLine, loginfo, row, isFullSweep, ref isFirst);
                    loginfo.Logs.Add(log);
                }
            }
            else
            {
                loginfo.Logs.Add(log);
            }
        }

        private static void UpdateTrimState(BenchLogItem log, BenchLogFile loginfo, bool isFullSweep, ref bool isTrim)
        {
            if (Regex.IsMatch(log.Operation, "Write_Trim", RegexOptions.IgnoreCase))
            {
                isTrim = true;
            }

            if (Regex.IsMatch(log.Operation, "Read_Cap", RegexOptions.IgnoreCase))
            {
                if (isTrim)
                {
                    if (isFullSweep)
                    {
                        loginfo.IsReadCapTrim = true;
                    }
                }
            }
            if (Regex.IsMatch(log.Operation, "Trim_End", RegexOptions.IgnoreCase))
            {
                isTrim = false;
            }
        }

        private static void IsSwapReedCap(BenchLogItem currentlog)
        {
            currentlog.Capinfos = currentlog.Capinfos
            .OrderBy(c => c.MSB)
            .ToList();
        }
        private static bool IsNeedMergeReadCap(BenchLogItem currentLog, BenchLogItem lastLog)
        {
            if (!lastLog.Operation.Equals("read_cap", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!currentLog.Address.Equals(lastLog.Address, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (currentLog.Capinfos.Exists
                (p => (lastLog.MSB) <= p.MSB && lastLog.MSB >= p.LSB ||
                (lastLog.LSB <= p.MSB && lastLog.LSB >= p.LSB)))
            {
                return false;
            }

            return true;
        }

        private static bool IsNeedMergeWriteSrc(BenchLogItem currentLog, BenchLogItem lastLog)
        {
            if (!lastLog.Operation.Equals("write_src", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!currentLog.Address.Equals(lastLog.Address, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true;
        }

        public static BenchLogFile GetLogPatternName(string originpath, bool isFullSweep)
        {
            var loginfo = new BenchLogFile();
            loginfo.Offset = -1;
            string line = "";
            int i = 0;
            bool isNeedDelete = false;
            string tmp = originpath;
            bool isTrim = false;
            while (IsFileLock(ref tmp))
            {
                if (!File.Exists(tmp))
                {
                    File.Copy(originpath, tmp);
                    isNeedDelete = true;
                }
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
                        loginfo.GetHeader(line);
                    }
                    var log = new BenchLogItem(line, loginfo, row, isFullSweep, ref isFirst);
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
                        isTrim = false;
                        loginfo.PatternSubName = log.RawData.Split(',')[3];
                        loginfo.Type = log.RawData.Split(',')[2];
                        loginfo.Interface = log.RawData.Split(',')[1];
                        loginfo.Version = string.IsNullOrEmpty(log.RawData.Split(',')[4]) ? "1" : log.RawData.Split(',')[4];
                        loginfo.DateCode = string.IsNullOrEmpty(log.RawData.Split(',')[5]) ? "" : log.RawData.Split(',')[5];
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

                    if (Regex.IsMatch(log.Operation, "Write_Trim", RegexOptions.IgnoreCase))
                    {
                        isTrim = true;
                    }

                    if (Regex.IsMatch(log.Operation, "Read_Cap", RegexOptions.IgnoreCase))
                    {
                        if (isTrim)
                        {
                            if (isFullSweep)
                            {
                                loginfo.IsReadCapTrim = true;
                            }
                        }
                    }
                    if (Regex.IsMatch(log.Operation, "Trim_End", RegexOptions.IgnoreCase))
                    {
                        isTrim = false;
                    }
                }
                catch (Exception)
                {
                    ;
                }
                i++;
            }
            sr.Close();
            if (isNeedDelete)
            {
                File.Delete(tmp);
            }

            return loginfo;
        }


    }

    //WRITE,JTAG,0x4400c01c,4,4,bennutc_wl01.regs.bennutc.pd_radio.common_sys.XTAL.top_xtal_reg_7.top_xtal_dctp_en_1p0_da,0x00000001,,
    [Serializable]
    public class Register
    {
        public Dictionary<string, List<string>> SrcDic = new Dictionary<string, List<string>>();
        public Dictionary<string, string> CapDic = new Dictionary<string, string>();
        public string Address;
        public string Length;
        public string Data = "";

        public static string GetAssignReg(string register, List<string> srcHistory, BenchLogItem log)
        {
            if (log.RawData.ToLower().IndexOf("write_trim") >= 0)
            {
                return register;
            }
            else
            {
                int index = srcHistory.FindIndex(p => p.Equals(srcHistory.Last()));
                if (index == 0)
                {
                    return register;
                }
                else
                {
                    return $"{register}_{index}";
                }
            }
        }

        public bool IsNeedSouce
        {
            get { return Data.Any(p => p.Equals('D')); }
        }

        public List<string> GetFieldList(BenchLogItem log)
        {
            var nameList = new List<string>();
            var result = Enumerable.Range(0, int.Parse(Length)).Select(p => $"[{p}]").ToList();
            string regMSBLSB = @"M(?<msb>\d+)_L(?<lsb>\d+)";

            long addrQuot = Convert.ToInt32(Address, 16) / 4;
            bool isNeedShift = Length == "64" && (addrQuot % 2 != 0);

            foreach (KeyValuePair<string, List<string>> src in SrcDic)
            {
                //check overwrite register
                string assignReg = GetAssignReg(src.Key, src.Value, log);

                Match regSrcKey = Regex.Match(src.Key, regMSBLSB, RegexOptions.IgnoreCase);
                int msb = int.Parse(regSrcKey.Groups["msb"].Value);
                int lsb = int.Parse(regSrcKey.Groups["lsb"].Value);
                if (isNeedShift)
                {
                    msb += 32;
                    lsb += 32;
                }

                string currentData = Data.Substring(int.Parse(Length) - 1 - msb, msb - lsb + 1);
                if (currentData.Any(p => p == 'D'))
                {
                    for (int i = lsb; i <= msb; i++)
                    {
                        result[i] = $"{assignReg}[{i - lsb}]";
                    }
                }
            }

            foreach (string cap in CapDic.Keys)
            {
                Match regCap = Regex.Match(cap, regMSBLSB, RegexOptions.IgnoreCase);
                int msb = int.Parse(regCap.Groups["msb"].Value);
                int lsb = int.Parse(regCap.Groups["lsb"].Value);
                if (isNeedShift)
                {
                    msb += 32;
                    lsb += 32;
                }
                for (int i = lsb; i <= msb; i++)
                {
                    result[i] = $"{cap}[{i - lsb}]";
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

        public string GetVectorInfo()
        {
            return $"// {CurrentTime} V:{VectorCount} C:{RealCount}";
        }

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
        public NamingBox(BenchLogFile file, string patternName)
        {
            PatternName = patternName;
            PatternFullName = patternName;
            SrmVecName = file.Logs.Exists(p => Regex.IsMatch(p.Operation, "meas", RegexOptions.IgnoreCase) ||
                (Regex.IsMatch(p.Operation, "autogen", RegexOptions.IgnoreCase) && Regex.IsMatch(p.Operation, "TE_OPERATION", RegexOptions.IgnoreCase))
                )
                ? PatternName + "_srm_meas" : patternName + "_srm";
            var subrPostFixed = new List<string>();
            if (file.Logs.Exists(p => Regex.IsMatch(p.Operation, "read_cap|spac_mem_read", RegexOptions.IgnoreCase)))
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

    public class CapTrimInfo
    {
        public string TrimTarget;
        public string BestCodeFunction;
        public string LowLimit;
        public string HighLimit;
        public Dictionary<string, int> TrimBits = new Dictionary<string, int>();
        public Dictionary<string, List<int>> DataInfo = new Dictionary<string, List<int>>();
        public List<List<string>> AllCalcStoreNames = new List<List<string>>();

        public void SetTrim(BenchLogItem log)
        {
            if (!DataInfo.ContainsKey(log.FuseName))
            {
                DataInfo.Add(log.FuseName, new List<int>());
            }

            DataInfo[log.FuseName].Add(int.Parse(log.FieldVal));
            if (!TrimBits.ContainsKey(log.FuseName))
            {
                TrimBits.Add(log.FuseName, log.MSB - log.LSB + 1);
            }
        }

        public string GetTrimCalcInfo()
        {
            List<string> realcsname = new List<string> { };

            foreach (List<string> csn in AllCalcStoreNames)
            {
                if (csn.Count() == 0)
                {
                    continue;
                }

                if (csn.Count() >= 2 &&
                (BestCodeFunction.IndexOf("Seq1PlusSeq2", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 BestCodeFunction.IndexOf("Parallel2", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    realcsname.Add(csn[csn.Count() - 2]);
                }

                realcsname.Add(csn.Last());
            }

            //Calc_BestCode_CPP
            var result = new List<string>
            {
                TrimTarget,
                GetSweepRange(),
                BestCodeFunction,
                string.Join(";", DataInfo.Keys),
                string.Join(";", realcsname),
            };
            return string.Format("Calc_BestCode_CPP({0})", string.Join(",", result));
        }
        private string GetSweepRange()
        {
            var result = new List<string>();
            int maxcount = DataInfo.Last().Value.Count();
            int totalcode = 1;
            for (int i = DataInfo.Values.Count() - 1; i >= 0; i--)
            {
                totalcode *= (DataInfo.Values.ElementAt(i).Max() + 1);
            }

            if (maxcount == totalcode)
            {
                foreach (List<int> item in DataInfo.Values)
                {
                    result.Add(string.Format("{0}::{1}", item.First(), item.Last()));
                }

                return string.Join(";", result);
            }
            else
            {
                int codenum = 0;
                while (codenum < maxcount)
                {
                    List<int> singcodes = new List<int>();
                    foreach (List<int> dval in DataInfo.Values)
                    {
                        singcodes.Add(dval[codenum]);
                    }

                    result.Add(string.Format("[{0}]", string.Join(";", singcodes)));

                    codenum++;
                }
                return string.Join("#", result);
            }
        }
    }
}

