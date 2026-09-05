using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using TestPlanLib.Basic;
using TestPlanLib.PatternListCsvFile;


namespace Automation.GenerateIgxl.Basic.Business.GenPatSet.Business
{
    public class PatSetStatus
    {
        public string Used { get; set; }
        public string SubrsCnt { get; set; }
        public string Existed { get; set; }
        public string ValidTimeSet { get; set; }
        public string ContainSubr { get; set; }
        public string EfuseRepeatCount { get; set; }
        public string GenericPatName { get; }
        public string PatternName { get; set; }
        public string PatFileValue { get; set; }
        public string TimeDomain { get; set; }
        public string PatVmVector { get; set; }
        public bool IsNotFoundInHardipInfo { get; set; }
        public List<string> PatSubrVector { get; set; }
        public string PatternComment { get; set; }

        public bool SubrNoVm { get; set; }

        public PatSetStatus(string genericPatName, string fullPatternName)
        {
            GenericPatName = genericPatName;
            PatternName = fullPatternName;
            Used = "NoCheck";//1st check rule
            ValidTimeSet = "NoCheck";//2nd check rule
            Existed = "NoCheck";//3rd check rule
            ContainSubr = "NoCheck"; // 4th check rule
            PatFileValue = "N/A";
            PatVmVector = "N/A";
            TimeDomain = "";
            PatSubrVector = new List<string>();
            IsNotFoundInHardipInfo = false;
        }

        public bool IsVmVectorHaveSamePatName
        {
            get
            {
                if (PatVmVector == "N/A")
                {
                    return true;  // bypass NA item
                }


                if (PatternName.Split('.').First().Equals(PatVmVector, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                return false;
            }
        }
    }

    public class PatSetGenerator
    {
        #region Field
        private readonly bool _skipPatternExistCheck;
        private readonly EnumInstrument _isUltraFlexPlusPattern;
        private string _patFolderDir;
        private readonly List<PatSetStatus> _allPatSetStatus = new List<PatSetStatus>();
        private PatSetSheet _patSetSheetAllnoPat;
        private PatSetSubSheet _patSubrSheetAllnoPat;
        private Dictionary<string, string> _fileListinAllFolder = new Dictionary<string, string>();
        private Dictionary<string, HashSet<string>> _fileInfolderMapping;
        private Dictionary<string, HardIpInfo> _hardIpInfoAlldic = new Dictionary<string, HardIpInfo>();
        private Dictionary<string, List<HardIpInfo>> _hardIpInfoGroupVersion = new Dictionary<string, List<HardIpInfo>>();
        private readonly Dictionary<string, string> _ufpfileListinAllFolder = new Dictionary<string, string>();
        private static readonly Regex _regex = new Regex("gz:|pat:|patx:", RegexOptions.IgnoreCase);
        #endregion

        #region Const
        private const string PatSetSheetName = "PatSets_All";
        private const string PatSubrSheetName = "Pattern_Subroutine";
        private const string DefaultBurst = "NO";
        private const string UsedType = "USE";
        private const string DebugType = "DEBUG";
        [SuppressMessage("Design", "Ter402", Justification = "IGXL output requires Windows-style backslash paths")]
        private const string PatPathPrefix = @".\Pattern\";
        #endregion

        #region Property
        private string PatExtensionSuffix
        {
            get
            {
                return _isUltraFlexPlusPattern == EnumInstrument.UFPlus ? ".PATX.GZ" : ".PAT.GZ";
            }

        }
        private string PatSuffix
        {
            get
            {
                return _isUltraFlexPlusPattern == EnumInstrument.UFPlus ? ".PATX" : ".PAT";
            }
        }
        public PatSetSheet PatSetSheetAll { get; set; }
        public PatSetSubSheet PatSubrSheetAll { get; set; }
        #endregion

        #region Constructor
        public PatSetGenerator(string patFolder)
        {
            _patFolderDir = patFolder;//LocalSpecs.TSMC_ENV ? patFolder : LocalSpecs.PatternPath;
            _skipPatternExistCheck = LocalSpecs.Options.IsIgnorePatternCheck;// ProjectConfigSingleton.Instance().GetConfigSettingValue("Basic", "IsIgnorePatternCheck").Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            _isUltraFlexPlusPattern = LocalSpecs.Options.InstrumentType;
        }
        #endregion

        #region Method
        public void GenerateFlow(List<PatternData> patternList)
        {
            GetPatFileDirAndPatInfo();

            GenPatSet(patternList);
        }

        public void GenerateFlow(List<PatternData> patternList, List<HardIpInfo> hardIpInfoAlldic, Dictionary<string, string> fileListinAllFolder)
        {
            _hardIpInfoAlldic = hardIpInfoAlldic.ToDictionary(x => x.FullPattern.ToUpper(), x => x);
            _hardIpInfoGroupVersion = LocalSpecs.HardIpInfos;

            _fileListinAllFolder = fileListinAllFolder;

            GenPatSet(patternList);
        }

        private void GenPatSet(List<PatternData> patternList)
        {
            PatSetSheetAll = new PatSetSheet(PatSetSheetName);
            PatSubrSheetAll = new PatSetSubSheet(PatSubrSheetName);
            _patSetSheetAllnoPat = new PatSetSheet("patSetSheetAllnoPat"); //temporary storage
            _patSubrSheetAllnoPat = new PatSetSubSheet("patSubrSheetAllnoPat"); //temporary storage
            _fileInfolderMapping = _fileListinAllFolder.Keys.GroupBy(x => Path.GetFileName(x).Split('.')[0]).ToDictionary(file => file.Key, file => file.ToHashSet(), StringComparer.OrdinalIgnoreCase);

            foreach (PatternData pattern in patternList)
            {
                if (pattern.Use.ToUpper().Equals(UsedType) || pattern.Use.ToUpper().Equals(DebugType))
                {
                    PatSetSheet patSetSheetAll = PatSetSheetAll;
                    PatSetSubSheet patSubrSheetAll = PatSubrSheetAll;
                    GenPatsetData(pattern, _fileListinAllFolder, ref patSetSheetAll, ref patSubrSheetAll);
                }
                else
                {
                    var patSetStatus = new PatSetStatus(pattern.PatternName, GetFileName(pattern.FileVersion))
                    {
                        Used = "Don't Use"
                    };
                    _allPatSetStatus.Add(patSetStatus);
                }
            }

            foreach (PatSet item in _patSetSheetAllnoPat.Rows)
            {
                item.IsBackup = true;
                item.PatSetRows.ForEach(x => x.IsBackup = true);

                PatSetSheetAll.AddRow(item);
            }


            if (_patSubrSheetAllnoPat.Rows.Any())
            {
                foreach (PatSetSubRow item in _patSubrSheetAllnoPat.Rows)
                {
                    item.IsBackup = true;
                    PatSubrSheetAll.Rows.Add(item);
                }
            }
            UpdatePatternListStatus();
            // Mod by JN if not subrount row igxl will meet error
            if (PatSubrSheetAll.Rows.Count == 0)
            {
                PatSubrSheetAll = null;
            }

            GenMissingPatErrorReport();
        }

        private void UpdatePatternListStatus()
        {
            Dictionary<string, PatSet> patsetRowDic = PatSetSheetAll.PatSetRowDic;
            foreach (KeyValuePair<string, PatternData> pat in AcTSetCategoryMapSingleton.Instance().PatternList)
            {
                if (pat.Value.IsExist)
                {
                    continue;
                }

                PatSet existPatset = patsetRowDic.TryGetValue(pat.Key, out PatSet value) ? value : null;
                if (existPatset != null && !existPatset.IsBackup)
                {
                    pat.Value.IsExist = true;
                }
            }
        }

        private void GenMissingPatErrorReport()
        {
            IEnumerable<string> missingPatterns = _patSetSheetAllnoPat.Rows.Select(x => x.PatSetName).Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (string missingPattern in missingPatterns)
            {
                ErrorReportManager.AddError(PatternMissing.E_MissingPatternFile_01, "", 1, 0, "{0} doesn't exist in pattern folder.", new string[] { missingPattern }, new ErrorInfo() { Pattern = missingPattern });
            }
        }

        [SuppressMessage("Design", "Ter403", Justification = "IGXL output requires Windows-style backslash paths")]
        private GenPatsetObj ResolvePatternPathAndStatus(GenPatsetObj genPatsetObj, PatternData pattern, PatSetStatus patSetStatus, Dictionary<string, string> fileListinAllFolder, ref string patternName)
        {
            if (_skipPatternExistCheck)
            { // don't link pattern with server pattern data
                if (LocalSpecs.CompileItem != null)
                {
                    if (LocalSpecs.CompileItem.TryGetValue(patternName, out CompileItem value))
                    {
                        genPatsetObj.FileValue = PatPathPrefix + value.TpCategory + @"\" + genPatsetObj.FileValue;
                        patSetStatus.Existed = "Skipped";
                    }
                    else
                    {
                        var patternObj = new PatternNameInfo(patternName);
                        genPatsetObj.FileValue = PatPathPrefix + patternObj.TpCategory + @"\" + genPatsetObj.FileValue;
                        patSetStatus.Existed = "NonExisted";
                        pattern.IsExist = false;
                    }
                }
            }
            else if (SetExistPatternInfo(patternName, ref genPatsetObj, fileListinAllFolder))
            {
                patSetStatus.Existed = "Existed";
            }
            //This condition is try to select max pattern version from HARDIPInfo to get pattern info
            else if (_hardIpInfoGroupVersion.Any() && LocalSpecs.SearchValidPatternRev)
            {
                _hardIpInfoGroupVersion.TryGetValue(pattern.PatternName.ToUpper(), out List<HardIpInfo> patInfos);
                if (patInfos != null)
                {
                    var versionList = patInfos.Select(x => int.Parse(x.Version.Split('_')[0])).ToList();
                    var sortVersionList = versionList.OrderByDescending(x => x).ToList();
                    foreach (int version in sortVersionList)
                    {
                        int versionPosition = versionList.IndexOf(version);
                        HardIpInfo versionPatternInfo = patInfos[versionPosition];
                        string currentVersionPatternName = $"{versionPatternInfo.Payload}_{versionPatternInfo.Version}".ToUpper();
                        //Check folder pattern again.
                        if (SetExistPatternInfo(currentVersionPatternName, ref genPatsetObj, fileListinAllFolder))
                        {
                            patSetStatus.Existed = "Existed";
                            patSetStatus.PatternComment = $"Pattern's version changes from {patternName} to {currentVersionPatternName}";
                            ErrorReportManager.AddError(PatternMissing.E_VersionChange_01, "", 1, 0, "Pattern's version changes from {0} to {1}", new string[] { patternName, currentVersionPatternName }, new ErrorInfo() { Pattern = patternName });
                            //Assign pattern name to find result pattern
                            patternName = currentVersionPatternName;
                            break;
                        }
                    }
                }
                if (patSetStatus.Existed != "Existed")
                {
                    var patternObj = new PatternNameInfo(patternName);
                    genPatsetObj.FileValue = PatPathPrefix + patternObj.TpCategory + @"\" + genPatsetObj.FileValue;
                    patSetStatus.Existed = "NonExisted";
                    pattern.IsExist = false;
                }
            }
            else
            {
                // create fake directory
                var patternObj = new PatternNameInfo(patternName);
                genPatsetObj.FileValue = PatPathPrefix + patternObj.TpCategory + @"\" + genPatsetObj.FileValue;
                patSetStatus.Existed = "NonExisted";
                pattern.IsExist = false;
            }

            return genPatsetObj;
        }

        private GenPatsetObj PopulatePatternExecutionMetadata(
            GenPatsetObj genPatsetObj,
            PatSetStatus patSetStatus,
            List<string> patSetSubFileValues,
            string patternName
        )
        {
            if (_hardIpInfoAlldic.TryGetValue(patternName.ToUpper(), out HardIpInfo reference))
            {
                reference.UseThisVersion = true;
                if (reference.SubrList != null)
                {
                    genPatsetObj.ContainSubr = true;
                    patSetStatus.ContainSubr = "ContainSubr";
                    foreach (string subr in reference.SubrList)
                    {
                        string patSetSubFileValue = genPatsetObj.FileValue + ":" + subr;
                        patSetStatus.PatSubrVector.Add(subr);
                        patSetSubFileValues.Add(patSetSubFileValue.ToUpper());
                    }
                }
                if (reference.TimeDomain != null)
                {
                    patSetStatus.TimeDomain = reference.TimeDomain;
                }
                string vmVector = reference.VmVector;
                if (!string.IsNullOrEmpty(vmVector))
                {
                    genPatsetObj.FileValue = genPatsetObj.FileValue + ":" + vmVector;
                    patSetStatus.PatVmVector = vmVector;
                }
                else if (genPatsetObj.ContainSubr)
                {
                    patSetStatus.SubrNoVm = true;
                }

                string subrsCnt = reference.CallSubRoutinesCntDiff;
                if (!string.IsNullOrEmpty(subrsCnt))
                {
                    patSetStatus.SubrsCnt = subrsCnt;
                }

                string efuseRepeatCount = reference.EfuseRepeatCount;
                if (!string.IsNullOrEmpty(efuseRepeatCount))
                {
                    patSetStatus.EfuseRepeatCount = efuseRepeatCount;
                }

            }
            return genPatsetObj;
        }

        private bool SetExistPatternInfo(string patternName, ref GenPatsetObj genPatsetObj, Dictionary<string, string> fileListinAllFolder)
        {
            if (!_fileInfolderMapping.TryGetValue(patternName, out HashSet<string> values))
            {
                return false;
            }
            string foundPatternName = values.FirstOrDefault().ToUpper();
            genPatsetObj.FileValue = fileListinAllFolder[foundPatternName].ToUpper();
            _patFolderDir = _patFolderDir.ToUpper();
            genPatsetObj.FileValue = genPatsetObj.FileValue.Replace(_patFolderDir, PatPathPrefix).Replace(@"\\", @"\").Replace("/", "\\");
            return true;
        }

        private class GenPatsetObj
        {
            public string FileValue { get; set; }

            public bool ContainSubr { get; set; }
        }

        private void GenPatsetData(
            PatternData pattern,
            Dictionary<string, string> fileListinAllFolder,
            ref PatSetSheet patsetsheet,
            ref PatSetSubSheet patsetsub
        )
        {
            try
            {
                bool isUfpOnly = !TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);
                //string fileValue = GetFileName(pattern.FileVersion, isgenufp).ToUpper();
                var patsetObj = new GenPatsetObj
                {
                    FileValue = GetFileName(pattern.FileVersion, isUfpOnly).ToUpper(),
                    ContainSubr = false
                };
                var patSetStatus = new PatSetStatus(pattern.PatternName.ToUpper(), patsetObj.FileValue) { Used = "Used" };

                if (IsNaPattern(pattern, patSetStatus))
                {
                    return;
                }

                string patternFileBaseName = patsetObj.FileValue
                    .Trim()
                    .Replace('/', '\\')
                    .TrimEnd('\\')
                    .ToUpper();
                patSetStatus.ValidTimeSet = "VailidTs";
                string patternName = patternFileBaseName.Replace(isUfpOnly ? ".PATX.GZ" : PatExtensionSuffix, "");
                //bool containSubr = false;
                patSetStatus.SubrNoVm = false;  // only sub pattern
                var patSetSubFileValues = new List<string>();
                patsetObj = ResolvePatternPathAndStatus(
                    patsetObj,
                    pattern,
                    patSetStatus,
                    fileListinAllFolder,
                    ref patternName
               );
                patsetObj = PopulatePatternExecutionMetadata(
                    patsetObj,
                    patSetStatus,
                    patSetSubFileValues,
                    patternName
                );

                if (!_regex.IsMatch(patsetObj.FileValue))
                {
                    patsetObj.FileValue = patsetObj.FileValue + ":" + patternName;
                    // always put vm_vector name no matter subroutine existed or not. 2017/05/05 Osprey Team.
                }
                patsetObj.FileValue = patsetObj.FileValue.Replace(_patFolderDir, PatPathPrefix).Replace(@"\\", @"\").Replace("/", "\\");
                if (!patsetObj.ContainSubr)
                {
                    patSetStatus.ContainSubr = "NonContainSubr";
                }

                patSetStatus.PatFileValue = patsetObj.FileValue;
                _allPatSetStatus.Add(patSetStatus);
                PatSet patSetItem = GenPatSetRow(patSetStatus);
                // Cyprus Rtos Subr pattern had been listed in Pattern List Csv 
                // is only Subr don't need to crate PatSetRow
                if (!patSetStatus.SubrNoVm)
                {
                    if (patSetStatus.Existed == "Existed" || patSetStatus.Existed == "Skipped")
                    {
                        if (string.Equals(pattern.Check, "FAIL", StringComparison.OrdinalIgnoreCase))
                        {
                            patSetItem.PatSetRows[0].Comment = pattern.CheckComment;
                            _patSetSheetAllnoPat.AddRow(patSetItem);
                        }
                        else
                        {
                            patsetsheet.AddRow(patSetItem);
                        }
                    }
                    else
                    {
                        _patSetSheetAllnoPat.AddRow(patSetItem);
                    }
                }

                if (!patsetObj.ContainSubr)
                {
                    return;
                }


                // add PatSubr
                List<PatSetSubRow> patSetSubRowList = GenPatSetSubRows(patSetSubFileValues, patSetStatus);
                foreach (PatSetSubRow patSetSubRow in patSetSubRowList)
                {
                    if (patSetStatus.Existed == "Existed" || patSetStatus.Existed == "Skipped")
                    {
                        if (string.Equals(pattern.Check, "FAIL", StringComparison.OrdinalIgnoreCase))
                        {
                            patSetSubRow.Comment = pattern.CheckComment;
                            _patSubrSheetAllnoPat.Rows.Add(patSetSubRow);
                        }
                        else
                        {
                            patsetsub.Rows.Add(patSetSubRow);
                        }
                    }
                    else
                    {
                        _patSubrSheetAllnoPat.Rows.Add(patSetSubRow);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                throw;
            }
        }

        private List<PatSetSubRow> GenPatSetSubRows(List<string> patSetSubFileValues, PatSetStatus patSetStatus)
        {
            var patSetSubRowList = new List<PatSetSubRow>();
            foreach (string patSetSubFileValue in patSetSubFileValues)
            {
                var patSetSubRow = new PatSetSubRow
                {
                    PatternFileName = patSetSubFileValue.ToUpper().Replace(".GZ", "")
                };
                if (patSetStatus.Existed == "NonExisted")
                {
                    patSetSubRow.Comment = "NonExisted";
                }

                if (patSetStatus.IsNotFoundInHardipInfo)
                {
                    patSetSubRow.ColumnA = "NotFoundInHardipInfo";
                }
                if (patSetStatus.SubrNoVm)
                {
                    patSetSubRow.ColumnA = "SubrNoVm";
                }

                patSetSubRowList.Add(patSetSubRow);
            }
            return patSetSubRowList;
        }

        private PatSet GenPatSetRow(PatSetStatus patSetStatus)
        {
            // Change PatSet toupper and remove .gz
            var patSet = new PatSetRow
            {
                Burst = DefaultBurst.ToUpper(),
                File = patSetStatus.PatFileValue.ToUpper().Replace(".GZ", ""),
                TimeDomain = patSetStatus.TimeDomain
            };

            var patSetItem = new PatSet { PatSetName = patSetStatus.GenericPatName };
            if (patSetStatus.Existed == "NonExisted")
            {
                patSet.Comment = "NonExisted";
            }

            if (!string.IsNullOrEmpty(patSetStatus.EfuseRepeatCount))
            {
                patSet.Comment = $"repeat:{patSetStatus.EfuseRepeatCount}";
            }
            if (!string.IsNullOrEmpty(patSetStatus.PatternComment))
            {
                patSet.Comment = patSetStatus.PatternComment;
            }

            patSetItem.AddRow(patSet);
            return patSetItem;
        }

        private bool IsNaPattern(PatternData pattern, PatSetStatus patSetStatus)
        {
            if (pattern.TimeSetVersion.Trim().ToUpper() == "NA" || pattern.FileVersion.Trim().ToUpper() == "NA" ||
                pattern.TimeSetVersion.Trim().ToUpper() == "N/A" || pattern.FileVersion.Trim().ToUpper() == "N/A")
            {
                patSetStatus.ValidTimeSet = "NoVailidTs";
                _allPatSetStatus.Add(patSetStatus);


                // add patSet for NA item
                var patSetNaItem = new PatSet { PatSetName = pattern.PatternName.ToUpper() };
                var patSetNa = new PatSetRow
                {
                    Burst = DefaultBurst.ToUpper(),
                    File = pattern.FileVersion.ToUpper().Replace(".GZ", ""),
                    Comment = "NonExisted-NA"
                };
                patSetNaItem.AddRow(patSetNa);

                _patSetSheetAllnoPat.AddRow(patSetNaItem);

                return true;
            }
            return false;
        }

        private void GetPatFileDirAndPatInfo()
        {
            if (!LocalSpecs.Options.IsIgnorePatternCheck)
            {
                if (_fileListinAllFolder.Count == 0)
                {
                    IEnumerable<string> files = AcTSetCategoryMapSingleton.Instance().PatternPathListInK.Where(s => s.EndsWith(PatSuffix, StringComparison.OrdinalIgnoreCase) || s.EndsWith(PatExtensionSuffix, StringComparison.OrdinalIgnoreCase));
                    foreach (string tmp in files)
                    {
                        string fileName = Path.GetFileName(tmp).ToUpper();
                        if (!_fileListinAllFolder.ContainsKey(fileName))
                        {
                            _fileListinAllFolder.Add(fileName, tmp);
                        }
                    }
                }

                if (_isUltraFlexPlusPattern == EnumInstrument.UFlex_UFPlus &&
                    _ufpfileListinAllFolder.Count == 0)
                {
                    IEnumerable<string> filesufp = AcTSetCategoryMapSingleton.Instance().PatternPathListInK.Where(s => s.EndsWith(".PATX", StringComparison.OrdinalIgnoreCase) ||
                                                                                                                       s.EndsWith(".PATX.GZ", StringComparison.OrdinalIgnoreCase));
                    foreach (string tmp in filesufp)
                    {
                        string fileName = Path.GetFileName(tmp).ToUpper();
                        if (!_ufpfileListinAllFolder.ContainsKey(fileName))
                        {
                            _ufpfileListinAllFolder.Add(fileName, tmp);
                        }
                    }
                }
            }

            if (_hardIpInfoAlldic.Count == 0 && LocalSpecs.HardIpInfos.Count > 0 && LocalSpecs.HardIpInfos != null)
            {
                _hardIpInfoGroupVersion = LocalSpecs.HardIpInfos;
                _hardIpInfoAlldic = LocalSpecs.HardIpInfos.SelectMany(x => x.Value).ToDictionary(x => x.FullPattern, x => x, StringComparer.OrdinalIgnoreCase);
            }
        }

        private string GetFileName(string fileVersion, bool isgenUfp = false)
        {
            if (fileVersion.Trim().ToUpper() == "NA" || fileVersion.Trim().ToUpper() == "N/A")
            {
                return string.Empty;
            }

            fileVersion = fileVersion.Replace('/', '\\');
            int startIndex = fileVersion.LastIndexOf('\\');
            string fileName = fileVersion;
            if (startIndex > 0)
            {
                fileName = fileVersion.Substring(startIndex + 1);
            }

            int endIndex = fileName.IndexOf('.');
            if (endIndex > 0)
            {
                fileName = fileName.Substring(0, endIndex);
                fileName += isgenUfp ? ".PATX.GZ" : PatExtensionSuffix;
            }

            return fileName;
        }

        public void GenPatSetErrorReport()
        {
            int rowNum = 0;
            foreach (PatSetStatus patSetStatus in _allPatSetStatus)
            {

                if (!patSetStatus.IsVmVectorHaveSamePatName)
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_01, "PatSetError", rowNum, 0, $"VM Vector name is mismatch, Pattern:{patSetStatus.PatternName}, VmVector:{patSetStatus.PatVmVector} ", new string[] { patSetStatus.PatternName, patSetStatus.PatVmVector });
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_02, "PatSetError", rowNum, 0, $"VM Vector name is mismatch, Pattern:{patSetStatus.PatternName}, VmVector:{patSetStatus.PatVmVector} ", new string[] { patSetStatus.PatternName, patSetStatus.PatVmVector });
                }

                rowNum++;
            }


        }

        #endregion
    }

    public class PatInfoCmd
    {
        // Get information from #.Pat
        public bool ConvertByArgs(string sourcePat, ref string output, string strArg)
        {

            string args = $" {strArg} \"{sourcePat}\"";
            try
            {
                string igxlRoot = Environment.GetEnvironmentVariable("igxlroot");
                if (string.IsNullOrEmpty(igxlRoot))
                {
                    return false;
                }

                var p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.FileName = Path.Combine(igxlRoot, "bin", "patinfo.exe");
                p.StartInfo.Arguments = args;
                p.Start();

                output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                if (p.ExitCode == 0)
                {
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                throw new Exception(e.StackTrace);
            }
        }

        public List<string> GetPinList(List<string> list)
        {
            //Case -1
            //Radix 4 for 1 - pin(s):
            //XO0(1 pin(s)) at offsets:  4068 		 Pin Names:  XO0 

            //Case -2
            //            Instrument Description Table:
            //-----------------------------
            //Name: DCVS, 434 bit(s), starting bit 36:
            //     Instrument Bitfield: 1
            //    CorePower1(14 pin(s)) at offsets:  36  67  98  129  160  191  222  253  284  315  346  377  408  439 		 Pin Names:  VDD_AFR  VDD_AVEMSR  VDD_DCS0  VDD_DCS1  VDD_DISP_EXT0  VDD_DISP_EXT1  VDD_DISP_INT  VDD_FABRIC  VDD_GPU0  VDD_GPU1  VDD_SOC  VDD_GPU_SRAM0  VDD_GPU_SRAM1  VDD_SRAM 

            var pinList = new List<string>();
            foreach (string line in list)
            {
                string context = line.Trim('\r').Trim();
                if (context == "")
                {
                    continue;
                }

                if (Regex.IsMatch(context, "Pin Names:", RegexOptions.IgnoreCase))
                {
                    int index = context.IndexOf("Pin Names:", StringComparison.Ordinal);
                    if (index != -1)
                    {
                        pinList.AddRange(context.Substring(index + 10).Trim().ToUpper().Split(' '));
                    }
                }
            }
            return pinList.Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
        }
    }
}
