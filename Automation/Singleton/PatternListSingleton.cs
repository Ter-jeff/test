using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.Static;

using LogLib.Utility;

using TestPlanLib.PatternListCsvFile;

namespace Automation.Singleton
{
    public class PatternListSingleton
    {

        private readonly string _oriPatListCsvFile;
        private readonly string _timeSetFolder;
        private readonly string _tempOriPatListCsvFile;
        private Dictionary<string, CompileItem> _compileList = new Dictionary<string, CompileItem>(StringComparer.CurrentCultureIgnoreCase);
        private readonly Dictionary<string, OriPatListItem> _oriPatCsvList = new Dictionary<string, OriPatListItem>(StringComparer.CurrentCultureIgnoreCase);        // c651 csv
        private static PatternListSingleton _instance;
        private bool _mergedWithCompileFile;
        private bool _updateLatestPat;

        private const string HeaderIndex = "#";
        private const string HeaderPattern = "Pattern";
        private const string HeaderLatestVersion = "Latest Version";
        private const string HeaderReleaseDate = "Release Date";
        private const string HeaderUseNoUse = "USE/No Use";
        private const string HeaderDri = "DRI";
        private const string HeaderReleaseNotes = "Release Notes";
        private const string HeaderRadar = "Radar #";
        private const string HeaderOrg = "Org";
        private const string HeaderTypeSpec = "Type Spec";
        private const string HeaderTimeSetVersion = "Timeset Version";
        private const string HeaderFileVersion = "File Versions";
        private const string HeaderOpcode = "OpCode";
        private const string HeaderScanMode = "ScanMode";
        private const string HeaderHalt = "Halt";
        private const string HeaderCompilation = "Compilation";
        private const string HeaderHlv = "HLV";
        private const string HeaderTpCategory = "T/P Category";
        private const string HeaderOriginTimeMode = "Original Timing Mode";
        private const string HeaderCheck = "Check";

        private const string HeaderRegularPattern = "#.*Pattern";
        private const string RegexHeader = @".*Pattern.*Timeset.*File\s+Versions.*";
        private static readonly Regex _regex = new Regex(".csv", RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(HeaderRegularPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"_\d+.TXT$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(@"(?<str>.*)_\d+.TXT$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex5 = new Regex(@".*_(?<ver>\d+).TXT$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex6 = new Regex("Timing Mode", RegexOptions.Compiled);
        private static readonly Regex _regex7 = new Regex(@",|\t|\s{4}", RegexOptions.Compiled);
        private static readonly Regex _regex8 = new Regex(",", RegexOptions.Compiled);
        private static readonly Regex _regex9 = new Regex(RegexHeader, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex10 = new Regex(@"^\s*\#\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex11 = new Regex("^Pattern$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex12 = new Regex(@"Latest\s+Version", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex13 = new Regex("USE.*No.*Use", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex14 = new Regex("DRI", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex15 = new Regex(@"Release\s+Date", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex16 = new Regex(@"Release\s+Note", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex17 = new Regex("Radar", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex18 = new Regex("Org", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex19 = new Regex(@"Type\s+Spec", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex20 = new Regex(@"Timeset\s+Latest", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex21 = new Regex(@"File\s+Versions", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex22 = new Regex(RegexHeader, RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex _regex23 = new Regex(@".*Pattern.*Release\s+Notes.*", RegexOptions.Compiled);
        private static readonly Regex _regex24 = new Regex("_DM_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex25 = new Regex("_SI_", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public bool HeaderFailFlag { get; private set; }

        public string CompiledPatternDashboardFile { get; set; }

        public Dictionary<string, OriPatListItem> PatternDictionary { get; } = new Dictionary<string, OriPatListItem>(StringComparer.CurrentCultureIgnoreCase);

        public static PatternListSingleton GetInstance(string filePath, string timeSetFolder)
        {
            return _instance ?? (_instance = new PatternListSingleton(filePath, timeSetFolder));
        }

        public static void Initialize()
        {
            _instance = null;
        }

        private PatternListSingleton(string filePath, string timeSetFolder)
        {
            _timeSetFolder = timeSetFolder;
            if (!File.Exists(filePath))
            //Judge if csv file exist in the folder
            {
                _oriPatCsvList = null;
            }
            else
            {
                _tempOriPatListCsvFile = _regex.Replace(filePath, "_DUM.csv");
                _oriPatListCsvFile = filePath;
                bool isTw = IsCompiledPatternDashboard();
                if (isTw)
                {
                    ReadTwCsvFile();
                }
                else
                {
                    ReadNonTwCsvFile();
                }
            }
        }

        private void ReadTwCsvFile()
        {
            Dictionary<string, int> headerIndexDic = new Dictionary<string, int>();

            if (File.Exists(_tempOriPatListCsvFile))
            {
                File.Delete(_tempOriPatListCsvFile);
            }
            File.Copy(_oriPatListCsvFile, _tempOriPatListCsvFile);
            var reader = new StreamReader(File.OpenRead(_tempOriPatListCsvFile));
            int cnt = 0;
            bool dataStart = false;
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine().Replace("\"", "");
                cnt++;
                bool blankRow = line.Replace(",", "").Trim() == "";
                string[] dataArray = line.Split(',');
                if (!dataStart && _regex2.IsMatch(line))
                {
                    //Read header
                    dataStart = true;
                    for (int i = 0; i < dataArray.Length; i++)
                    {
                        headerIndexDic.Add(dataArray[i].ToUpper(), i);
                    }

                }
                else if (dataStart)
                {
                    if (blankRow)
                    {
                        continue;
                    }
                    NormalizeTwDataArray(ref dataArray, headerIndexDic.Count);
                    //Read data
                    OriPatListItem oriPatCsvItem = BuildTwOriPatItem(dataArray, headerIndexDic, cnt);
                    if (!PatternDictionary.ContainsKey(oriPatCsvItem.Pattern.ToUpper()))
                    {
                        PatternDictionary.Add(oriPatCsvItem.Pattern.ToUpper(), oriPatCsvItem);
                    }
                    else
                    {
                        ErrorMessageBox.Show("Duplicated Pattern in List -> " + oriPatCsvItem.Pattern, "Pattern Duplicated");
                    }
                }
            }
            reader.Close();

            if (File.Exists(_tempOriPatListCsvFile))
            {
                File.Delete(_tempOriPatListCsvFile);
            }
        }

        internal static void NormalizeTwDataArray(ref string[] dataArray, int headerCount)
        {
            if (dataArray.Length < headerCount)
            {
                int originalNum = dataArray.Length;
                Array.Resize(ref dataArray, headerCount);
                string[] blankArray = Enumerable.Range(0, headerCount)
                    .Select(i => string.Empty)
                    .ToArray();
                Array.Copy(blankArray, originalNum, dataArray, originalNum, headerCount - originalNum);
            }
        }

        internal static string GetTwField(string[] dataArray, Dictionary<string, int> headerIndexDic, string header)
        {
            return headerIndexDic.ContainsKey(header.ToUpper()) ? dataArray[headerIndexDic[header.ToUpper()]] : "";
        }

        internal static OriPatListItem BuildTwOriPatItem(string[] dataArray, Dictionary<string, int> headerIndexDic, int cnt)
        {
            return new OriPatListItem
            {
                RowNum = cnt,
                Idx = GetTwField(dataArray, headerIndexDic, HeaderIndex),
                Pattern = GetTwField(dataArray, headerIndexDic, HeaderPattern),
                LatestVersion = GetTwField(dataArray, headerIndexDic, HeaderLatestVersion),
                ReleaseDate = GetTwField(dataArray, headerIndexDic, HeaderReleaseDate),
                UseNoUse = GetTwField(dataArray, headerIndexDic, HeaderUseNoUse),
                DRi = GetTwField(dataArray, headerIndexDic, HeaderDri),
                ReleaseNote = GetTwField(dataArray, headerIndexDic, HeaderReleaseNotes),
                RadarNum = GetTwField(dataArray, headerIndexDic, HeaderRadar),
                Org = GetTwField(dataArray, headerIndexDic, HeaderOrg),
                TypeSpec = GetTwField(dataArray, headerIndexDic, HeaderTypeSpec),
                TimeSetVersion = GetTwField(dataArray, headerIndexDic, HeaderTimeSetVersion),
                FileVersions = GetTwField(dataArray, headerIndexDic, HeaderFileVersion),
                OpCode = GetTwField(dataArray, headerIndexDic, HeaderOpcode),
                ScanMode = GetTwField(dataArray, headerIndexDic, HeaderScanMode),
                Halt = GetTwField(dataArray, headerIndexDic, HeaderHalt),
                Compilation = GetTwField(dataArray, headerIndexDic, HeaderCompilation),
                HLv = GetTwField(dataArray, headerIndexDic, HeaderHlv),
                TpCategory = GetTwField(dataArray, headerIndexDic, HeaderTpCategory),
                OriTimeMod = GetTwField(dataArray, headerIndexDic, HeaderOriginTimeMode),
                CheckRt = GetTwField(dataArray, headerIndexDic, HeaderCheck)
            };
        }

        public void ReadNonTwCsvFile()
        {
            #region CheckTiming Set
            Dictionary<string, TimeSetItem> dicTimeSetVersion = BuildTimeSetVersionDictionary();
            #endregion

            #region Pattern List Reader of Customer's format
            int idxIdx = -1;
            int idxPattern = -1;
            int idxLatestVersion = -1;
            int idxReleaseDate = -1;
            int idxUseNoUse = -1;
            int idxDRi = -1;
            int idxReleaseNote = -1;
            int idxRadarNum = -1;
            int idxOrg = -1;
            int idxTypeSpec = -1;
            int idxTimeSetLatest = -1;
            int idxFileVersions = -1;
            int idxOpCode = -1;
            int idxScanMode = -1;
            int idxHalt = -1;
            int idxCompilation = -1;
            int idxTpCategory = -1;
            int idxHLv = -1;
            int idxScanTSet = -1;
            int headerCnt = 0;
            if (File.Exists(_tempOriPatListCsvFile))
            {
                File.Delete(_tempOriPatListCsvFile);
            }
            File.Copy(_oriPatListCsvFile, _tempOriPatListCsvFile);
            var reader = new StreamReader(File.OpenRead(_tempOriPatListCsvFile));
            bool dataStart = false;

            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine().Replace("\"", "");
                int valueCnt = _regex8.Matches(line).Count + 1;
                if (headerCnt > 0)
                {
                    for (int iChr = 0; iChr < headerCnt - valueCnt; iChr++)
                    {
                        line += ",";
                    }
                }

                string[] values = line.Split(',');
                if (!dataStart && _regex9.IsMatch(line)) //the 1st line, Header
                {
                    headerCnt = values.Length;
                    for (int item = 0; item < values.Length; item++)
                    {
                        ParseNonTwHeaderColumn(values, item,
                            ref idxIdx, ref idxPattern, ref idxLatestVersion, ref idxReleaseDate,
                            ref idxUseNoUse, ref idxDRi, ref idxReleaseNote, ref idxRadarNum,
                            ref idxOrg, ref idxTypeSpec, ref idxTimeSetLatest, ref idxFileVersions,
                            ref idxOpCode, ref idxScanMode, ref idxHalt, ref idxCompilation,
                            ref idxTpCategory, ref idxScanTSet, ref idxHLv);
                    }
                    dataStart = true;
                    HeaderFailFlag = false;
                    continue;
                }
                if (!dataStart && !_regex22.IsMatch(line))
                {
                    HeaderFailFlag = true;
                }
                if (dataStart)
                {
                    if (values[idxPattern].Length > 10) //key
                    {
                        var oriPatCsvItem = new OriPatListItem();
                        PopulateNonTwOriPatItem(oriPatCsvItem, values, dicTimeSetVersion,
                            idxIdx, idxPattern, idxLatestVersion, idxReleaseDate, idxUseNoUse,
                            idxDRi, idxReleaseNote, idxRadarNum, idxOrg, idxTypeSpec,
                            idxTimeSetLatest, idxFileVersions, idxOpCode, idxScanMode, idxHalt,
                            idxCompilation, idxTpCategory, idxHLv, idxScanTSet);

                        string fileVersionPart = values[idxFileVersions].Trim().ToUpper() == "N/A" || values[idxFileVersions].Trim().ToUpper() == "NA"
                            ? "NA"
                            : values[idxFileVersions].Split('/').Last().ToUpper().Replace(".ATP", "").Replace(".GZ", "");
                        string patKeyName = values[idxPattern] + "#" + fileVersionPart;

                        string genericName = values[idxPattern].ToUpper();

                        if (!_oriPatCsvList.ContainsKey(patKeyName))
                        {
                            _oriPatCsvList.Add(patKeyName, oriPatCsvItem);
                            PatternDictionary.Add(genericName, oriPatCsvItem);
                        }
                        else
                        {
                            ErrorMessageBox.Show("Duplicated Pattern in List -> " + values[idxPattern], "Pattern Duplicated");
                        }
                    }
                }


            }
            reader.Close();
            if (File.Exists(_tempOriPatListCsvFile))
            {
                File.Delete(_tempOriPatListCsvFile);
            }

            #endregion
        }

        private Dictionary<string, TimeSetItem> BuildTimeSetVersionDictionary()
        {
            var timeSetFiles = new List<string>();
            if (!string.IsNullOrEmpty(_timeSetFolder) && Directory.Exists(_timeSetFolder))
            {
                timeSetFiles = Directory.GetFiles(_timeSetFolder, "TIMESET*.txt", SearchOption.TopDirectoryOnly).ToList();
            }

            var dicTimeSetVersion = new Dictionary<string, TimeSetItem>();
            foreach (string file in timeSetFiles)
            {
                string setName = Path.GetFileName(file);
                if (_regex3.IsMatch(setName))
                {
                    string timeSet = _regex4.Match(setName).Groups["str"].ToString().ToUpper();
                    int paraVer = Convert.ToInt32(_regex5.Match(setName).Groups["ver"].ToString());
                    var timeItem = new TimeSetItem { Version = paraVer };
                    ReadTimeSetTimeMode(file, timeItem);

                    if (!dicTimeSetVersion.ContainsKey(timeSet))
                    {
                        dicTimeSetVersion.Add(timeSet, timeItem);
                    }
                    else
                    {
                        if (dicTimeSetVersion[timeSet].Version < timeItem.Version)
                        {
                            dicTimeSetVersion[timeSet].Version = timeItem.Version;
                            dicTimeSetVersion[timeSet].TimeMod = timeItem.TimeMod;
                        }
                    }
                }
            }
            return dicTimeSetVersion;
        }

        private void ReadTimeSetTimeMode(string file, TimeSetItem timeItem)
        {
            var streamReader = new StreamReader(file);
            do
            {
                string line = streamReader.ReadLine();
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }

                if (_regex6.IsMatch(line))
                {
                    string[] tmpAry = _regex7.Split(line.Trim());
                    if (tmpAry.Length == 1)
                    {
                        string msg =
                            $"This timeSet : {Path.GetFileNameWithoutExtension(file)} header can not recognize by tool";
                        ErrorMessageBox.Show(msg, "Header not recognize");
                        continue;
                    }
                    timeItem.TimeMod = tmpAry[1].ToUpper();
                    break;

                }

            } while (streamReader.Peek() != -1);
            streamReader.Close();
        }

        private void ParseNonTwHeaderColumn(string[] values, int item,
            ref int idxIdx, ref int idxPattern, ref int idxLatestVersion, ref int idxReleaseDate,
            ref int idxUseNoUse, ref int idxDRi, ref int idxReleaseNote, ref int idxRadarNum,
            ref int idxOrg, ref int idxTypeSpec, ref int idxTimeSetLatest, ref int idxFileVersions,
            ref int idxOpCode, ref int idxScanMode, ref int idxHalt, ref int idxCompilation,
            ref int idxTpCategory, ref int idxScanTSet, ref int idxHLv)
        {
            if (_regex10.IsMatch(values[item]))
            {
                idxIdx = item;
            }

            if (_regex11.IsMatch(values[item]))
            {
                idxPattern = item;
            }

            if (_regex12.IsMatch(values[item]))
            {
                idxLatestVersion = item;
            }

            if (_regex13.IsMatch(values[item]))
            {
                idxUseNoUse = item;
            }

            if (_regex14.IsMatch(values[item]))
            {
                idxDRi = item;
            }

            if (_regex15.IsMatch(values[item]))
            {
                idxReleaseDate = item;
            }

            if (_regex16.IsMatch(values[item]))
            {
                idxReleaseNote = item;
            }

            if (_regex17.IsMatch(values[item]))
            {
                idxRadarNum = item;
            }

            if (_regex18.IsMatch(values[item]))
            {
                idxOrg = item;
            }

            if (_regex19.IsMatch(values[item]))
            {
                idxTypeSpec = item;
            }

            if (_regex20.IsMatch(values[item]))
            {
                idxTimeSetLatest = item;
            }

            if (_regex21.IsMatch(values[item]))
            {
                idxFileVersions = item;
            }

            if (values[item] == "OpCode")
            {
                idxOpCode = item;
            }

            if (values[item] == "ScanMode")
            {
                idxScanMode = item;
            }

            if (values[item] == "Halt")
            {
                idxHalt = item;
            }

            if (values[item] == "Compilation")
            {
                idxCompilation = item;
            }

            if (values[item] == "T/P Category")
            {
                idxTpCategory = item;
            }

            if (values[item] == "ScanSetupTSet")
            {
                idxScanTSet = item;
            }

            if (values[item] == "HLV")
            {
                idxHLv = item;
            }
        }

        private void PopulateNonTwOriPatItem(OriPatListItem oriPatCsvItem, string[] values,
            Dictionary<string, TimeSetItem> dicTimeSetVersion,
            int idxIdx, int idxPattern, int idxLatestVersion, int idxReleaseDate, int idxUseNoUse,
            int idxDRi, int idxReleaseNote, int idxRadarNum, int idxOrg, int idxTypeSpec,
            int idxTimeSetLatest, int idxFileVersions, int idxOpCode, int idxScanMode, int idxHalt,
            int idxCompilation, int idxTpCategory, int idxHLv, int idxScanTSet)
        {
            if (idxIdx >= 0)
            {
                oriPatCsvItem.Idx = values[idxIdx];
            }

            if (idxPattern >= 0)
            {
                oriPatCsvItem.Pattern = values[idxPattern];
            }

            if (idxLatestVersion >= 0)
            {
                oriPatCsvItem.LatestVersion = values[idxLatestVersion];
            }

            if (idxReleaseDate >= 0)
            {
                oriPatCsvItem.ReleaseDate = values[idxReleaseDate];
            }

            if (idxUseNoUse >= 0)
            {
                oriPatCsvItem.UseNoUse = values[idxUseNoUse];
            }

            if (idxDRi >= 0)
            {
                oriPatCsvItem.DRi = values[idxDRi];
            }

            if (idxReleaseNote >= 0)
            {
                oriPatCsvItem.ReleaseNote = values[idxReleaseNote];
            }

            if (idxRadarNum >= 0)
            {
                oriPatCsvItem.RadarNum = values[idxRadarNum];
            }

            if (idxOrg >= 0)
            {
                oriPatCsvItem.Org = values[idxOrg];
            }

            if (idxTypeSpec >= 0)
            {
                oriPatCsvItem.TypeSpec = values[idxTypeSpec];
            }

            if (idxTimeSetLatest >= 0)
            {
                oriPatCsvItem.TimeSetLatest = values[idxTimeSetLatest].ToUpper() == "N/A" ? "NA" : values[idxTimeSetLatest].Split('/').Last().ToUpper();
            }

            if (dicTimeSetVersion.ContainsKey(oriPatCsvItem.TimeSetLatest.Replace(".TXT", "")))
            {
                string key = oriPatCsvItem.TimeSetLatest.Replace(".TXT", "");
                oriPatCsvItem.TimeSetLatest = oriPatCsvItem.TimeSetLatest.Replace(".TXT", "") + "_" +
                                               dicTimeSetVersion[key].Version + ".TXT";
                oriPatCsvItem.OriTimeMod = dicTimeSetVersion[key].TimeMod.ToUpper();
            }

            if (idxFileVersions >= 0)
            {
                oriPatCsvItem.FileVersions = values[idxFileVersions];
            }

            if (idxOpCode >= 0)
            {
                oriPatCsvItem.OpCode = values[idxOpCode];
            }

            if (idxScanMode >= 0)
            {
                oriPatCsvItem.ScanMode = values[idxScanMode];
            }

            if (idxHalt >= 0)
            {
                oriPatCsvItem.Halt = values[idxHalt];
            }

            if (idxCompilation >= 0)
            {
                oriPatCsvItem.Compilation = values[idxCompilation];
            }

            if (idxTpCategory >= 0)
            {
                oriPatCsvItem.TpCategory = values[idxTpCategory];
            }

            if (idxHLv >= 0)
            {
                oriPatCsvItem.HLv = values[idxHLv];
            }

            if (idxScanTSet >= 0)
            {
                oriPatCsvItem.ScanTSet = values[idxScanTSet];
            }
        }

        public bool IsCompiledPatternDashboard()
        {
            if (_oriPatCsvList == null)
            {
                return true;
            }

            if (File.Exists(_tempOriPatListCsvFile))
            {
                File.Delete(_tempOriPatListCsvFile);
            }
            File.Copy(_oriPatListCsvFile, _tempOriPatListCsvFile);
            bool isTw = JudgeTw(_tempOriPatListCsvFile);
            if (File.Exists(_tempOriPatListCsvFile))
            {
                File.Delete(_tempOriPatListCsvFile);
            }
            return isTw;
        }

        private bool JudgeTw(string filePath)
        {
            #region Pattern List Reader of Customer's format
            int idxTpCategory = -1;
            int checkLine = 0;
            var reader = new StreamReader(File.OpenRead(filePath));
            bool dataStart = false;
            while (!reader.EndOfStream)
            {
                if (checkLine > 100)
                {
                    break;
                }

                checkLine++;
                string line = reader.ReadLine()?.Replace("\"", "");
                if (line != null)
                {
                    string[] values = line.Split(',');
                    if (dataStart)
                    {
                        if (values.Length < idxTpCategory)
                        {
                            reader.Close();
                            return false;
                        }
                        if (values[idxTpCategory].Trim() != "")
                        {
                            reader.Close();
                            return true;
                        }
                    }
                    if (_regex23.IsMatch(line)) //the 1st line
                    {
                        for (int item = 0; item < values.Length; item++)
                        {
                            if (values[item] == "T/P Category")
                            {
                                idxTpCategory = item;
                            }
                        }
                        if (idxTpCategory < 0)
                        {
                            reader.Close();
                            return false;
                        }
                        dataStart = true;
                    }
                }
            }
            reader.Close();
            return false;
            #endregion
        }

        public Dictionary<string, OriPatListItem> GetPatternData()
        {
            if (IsCompiledPatternDashboard())
            {
                return PatternDictionary;
            }
            return _oriPatCsvList;
        }

        public void UpdatePatternDashboardWithCompiledPatCsv(Dictionary<string, CompileItem> compileItems)
        {
            try
            {
                if (_oriPatCsvList == null)
                {
                    return;
                }
                if (IsCompiledPatternDashboard())
                {
                    //If Pattern List is TW File, directly return
                    CompiledPatternDashboardFile = _oriPatListCsvFile;
                    return;
                }
                if (_mergedWithCompileFile)
                {
                    return;
                }
                _compileList = compileItems;

                if (!_updateLatestPat)
                {
                    CompiledPatternDashboardFile = _oriPatListCsvFile.Replace(".csv", "_TW" + VersionControl.Timestamp.Replace(" ", "_") + ".csv");
                }
                else
                {
                    CompiledPatternDashboardFile = _oriPatListCsvFile.Replace(".csv", "_UseLatestPatternVersion_TW" + VersionControl.Timestamp.Replace(" ", "_") + ".csv");
                }

                #region Writing to TW format

                var csv = new StringBuilder();
                string newLine =
                    "#,Pattern,Latest Version,Release Date,USE/No Use,DRI,Release Notes,Radar #,Org,Type Spec,Timeset Version,File Versions,OpCode,ScanMode,Halt,Compilation,HLV,T/P Category,ScanSetupTSet,Original Timing Mode,Check,CheckComment";
                csv.AppendLine(newLine);

                foreach (string key in _oriPatCsvList.Keys)
                {
                    string generic = key.Split('#')[0].Replace("\"", "");
                    string realVersion = key.Split('#')[1].Replace("\"", "");

                    if (_oriPatCsvList[key].Idx == "")
                    {
                        newLine = "0";
                    }
                    else
                    {
                        newLine = _oriPatCsvList[key].Idx;
                    }

                    newLine += "," + _oriPatCsvList[key].Pattern;
                    newLine += "," + _oriPatCsvList[key].LatestVersion;
                    newLine += "," + _oriPatCsvList[key].ReleaseDate;
                    newLine += "," + _oriPatCsvList[key].UseNoUse;
                    newLine += "," + _oriPatCsvList[key].DRi;
                    newLine += "," + _oriPatCsvList[key].ReleaseNote;
                    newLine += "," + _oriPatCsvList[key].RadarNum;
                    newLine += "," + _oriPatCsvList[key].Org;
                    newLine += "," + _oriPatCsvList[key].TypeSpec;
                    newLine += "," + _oriPatCsvList[key].TimeSetLatest;
                    newLine += "," + _oriPatCsvList[key].FileVersions;

                    if (_compileList.ContainsKey(realVersion))
                    {
                        newLine += "," + _compileList[realVersion].OpCode;
                        _oriPatCsvList[key].OpCode = _compileList[realVersion].OpCode;

                        newLine += "," + _compileList[realVersion].ScanMode;
                        _oriPatCsvList[key].ScanMode = _compileList[realVersion].ScanMode;

                        newLine += "," + _compileList[realVersion].Halt;
                        _oriPatCsvList[key].Halt = _compileList[realVersion].Halt;

                        newLine += "," + _compileList[realVersion].Compilation;
                        _oriPatCsvList[key].Compilation = _compileList[realVersion].Compilation;

                        newLine += "," + _compileList[realVersion].HLv;
                        _oriPatCsvList[key].HLv = _compileList[realVersion].HLv;

                        newLine += "," + _compileList[realVersion].TpCategory;
                        _oriPatCsvList[key].TpCategory = _compileList[realVersion].TpCategory;
                        newLine += "," + _compileList[realVersion].ScanSetupTSet;
                        _oriPatCsvList[key].ScanTSet = _compileList[realVersion].ScanSetupTSet;
                        newLine += "," + _oriPatCsvList[key].OriTimeMod;
                        #region Validate TimingMode / Pattern Mode / Opcode Mode

                        string patNameOpcode = "";
                        if (_regex24.IsMatch(realVersion))
                        {
                            patNameOpcode = "DUAL";
                        }
                        else if (_regex25.IsMatch(realVersion))
                        {
                            patNameOpcode = "SINGLE";
                        }

                        string timingOpcode = _oriPatCsvList[key].OriTimeMod;
                        string patternOpcode = _compileList[realVersion].OpCode;

                        if (patternOpcode.ToUpper() != patNameOpcode)
                        {
                            newLine += "," + "Fail";
                            newLine += "," + "Pat Opcode / Pat Name Mode Mismatch";
                        }
                        else if (!string.Equals(patternOpcode, timingOpcode, StringComparison.OrdinalIgnoreCase))
                        {
                            newLine += "," + "Fail";
                            newLine += "," + $"Timing({timingOpcode}) / Pat({patternOpcode}) Opcode Mode Mismatch";
                        }
                        else
                        {
                            newLine += "," + "Pass";
                            newLine += ",";
                        }
                        #endregion
                    }
                    else
                    {
                        var pattern = new PatternNameInfo(generic);

                        newLine += "," + pattern.OpCode;
                        _oriPatCsvList[key].OpCode = pattern.OpCode;

                        newLine += "," + "NA";
                        _oriPatCsvList[key].ScanMode = "NA";

                        newLine += "," + "NA";
                        _oriPatCsvList[key].Halt = "NA";

                        newLine += "," + "NA";
                        _oriPatCsvList[key].Compilation = "NA";

                        newLine += "," + "NA";
                        _oriPatCsvList[key].HLv = "NA";

                        newLine += "," + pattern.TpCategory;
                        _oriPatCsvList[key].TpCategory = pattern.TpCategory;

                        newLine += "," + _oriPatCsvList[key].OriTimeMod;
                        newLine += "," + _oriPatCsvList[key].CheckRt;
                        newLine += ",";
                    }

                    csv.AppendLine(newLine);
                }
                //after your loop
                File.WriteAllText(CompiledPatternDashboardFile, csv.ToString());

                #endregion

                _mergedWithCompileFile = true;
            }
            catch (Exception e)
            {
                throw new Exception(e.StackTrace);
            }
        }

        public List<string> UpdateCsvByLatestPattern(Dictionary<string, string> latestPatternDict)
        {
            if (IsCompiledPatternDashboard())
            {
                return null;
            }

            if (_oriPatCsvList == null)
            {
                return null;
            }

            var modPatternList = new List<string>();

            foreach (string key in _oriPatCsvList.Keys)
            {
                string generic = key.Split('#')[0].Replace("\"", "");
                string realVersion = key.Split('#')[1].Replace("\"", "");
                if (_oriPatCsvList[key].FileVersions.Trim().ToUpper().Equals("N/A") || _oriPatCsvList[key].FileVersions.Trim().ToUpper().Equals("NA"))
                {
                    continue;
                }

                if (!latestPatternDict.TryGetValue(generic, out string value))
                {
                    continue;
                }

                string latestPatFile = generic + "_" + value;

                if (realVersion != latestPatFile)
                {
                    _updateLatestPat = true;

                    string modPattern = _oriPatCsvList[key].FileVersions.Replace(realVersion, latestPatFile);

                    if (modPattern.IndexOf(".", StringComparison.Ordinal) == -1)
                    {
                        modPattern += ".ATP.GZ";
                    }

                    _oriPatCsvList[key].FileVersions = modPattern;

                    modPatternList.Add("Update " + generic + " to latest Pattern version of Server: " + latestPatternDict[generic]);
                }
            }

            return modPatternList;
        }
    }
}
