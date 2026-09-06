using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Teradyne.Oasis.IGData;

using TestPlanLib.PatternListCsvFile;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class PatSetAllGenerator
    {
        #region Field and Property
        private const string PatSetSheetName = "PatSets_All";
        private const string PatSubrSheetName = "Pattern_Subroutine";
        private const string DefaultBurst = "NO";
        public const string UsedType = "USE";
        public const string DebugType = "DEBUG";
#pragma warning disable Ter402 // IGXL output requires Windows-style relative pattern path
        private const string PatPath = @".\Pattern\";
#pragma warning restore Ter402
        private const string PatSuffix = ".PAT.GZ";

        public Dictionary<string, PatSetStatus> LpatSetStatus = new Dictionary<string, PatSetStatus>();
        private readonly Dictionary<string, string> _dicPatPath;
        private PatSetSheet _patSetSheetAllnoPat;
        private PatSetSubSheet _patSubrSheetAllnoPat;
        private readonly Dictionary<string, string> _fileListinAllFolder = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public PatSetSheet PatSetSheetAll { get; set; }
        public PatSetSubSheet PatSubrSheetAll { get; set; }
        private readonly bool _genPatNoUsed;
        private string _patExtName;
        #endregion

        #region Constructor

        public PatSetAllGenerator(bool genPatNoUsed)
        {
            _genPatNoUsed = genPatNoUsed;
        }

        public PatSetAllGenerator(Dictionary<string, string> dicPatPath)
        {
            _dicPatPath = dicPatPath;
        }

        #endregion

        #region Methods
        public void GenerateFlow(string srcPath, string patternFolder, Dictionary<string, string> patterns, List<string> patternsInCharPlan, HashSet<string> noUsePatHash, HashSet<string> noExistPatHash, string patExtName = ".PAT")
        {
            PatSetSheetAll = new PatSetSheet(PatSetSheetName);
            PatSubrSheetAll = new PatSetSubSheet(PatSubrSheetName);
            _patSetSheetAllnoPat = new PatSetSheet("patSetSheetAllnoPat");//temporary storage
            _patSubrSheetAllnoPat = new PatSetSubSheet("patSubrSheetAllnoPat");//temporary storage
            _patExtName = patExtName;
            try
            {
                var patternsHashSet = LocalSpecs.PatternDatas.Values.Select(p => p.PatternName.ToUpper()).ToHashSet();
                foreach (string pattern in patternsInCharPlan)
                {
                    //if (LocalSpecs.PatternDatas.Values.FirstOrDefault(p => p.PatternName.ToUpper() == pattern.Split(':')[0].ToUpper()) == null)
                    if (!patternsHashSet.Contains(pattern.Split(':')[0].ToUpper()))
                    {
                        var patdata = new PatternData { PatternName = pattern.Split(':')[0], FileVersion = "N/A" };
                        LocalSpecs.PatternDatas.Add(pattern.Split(':')[0], patdata);
                    }
                }
            }
            catch (Exception)
            {
            }

            //PatternData
            foreach (PatternData pattern in LocalSpecs.PatternDatas.Values)
            {
                if (pattern.Use.ToUpper().Equals(UsedType) ||
                    _genPatNoUsed ||
                    LocalSpecs.PatternsInCharPlan.Contains(pattern.PatternName, StringComparer.InvariantCultureIgnoreCase))
                {
                    GenPatsetData(pattern, srcPath, patternFolder, patterns, noExistPatHash);
                }
                else
                {
                    var patSetStatus = new PatSetStatus { Used = "Don't Use" };
                    LpatSetStatus.Add(pattern.PatternName, patSetStatus);
                    noUsePatHash.Add(pattern.PatternName);
                }
            }

            foreach (PatSet item in _patSetSheetAllnoPat.Rows)
            {
                item.PatSetRows.ForEach(x => x.IsBackup = true);
                noExistPatHash.Add(item.PatSetName);
                PatSetSheetAll.AddRow(item);
            }

            foreach (PatSetSubRow item in _patSubrSheetAllnoPat.Rows.Where(item => PatSubrSheetAll != null))
            {
                PatSubrSheetAll.Rows.Add(item);
            }

            ModifyPatSetAll(PatSetSheetAll);
        }

        private void GenPatsetData(PatternData pattern, string srcPath, string patternFolder, Dictionary<string, string> patterns, HashSet<string> noExistPatHash)
        {
            string fileValue = CreateFileValue(pattern.FileVersion, out string baseFilePath);
            var patSetStatus = new PatSetStatus { Used = "Used" };
            if (srcPath == null || baseFilePath == null)
            {
                throw new Exception($"Directory object: {srcPath} or {baseFilePath} is null. ");
            }

            baseFilePath = baseFilePath.Trim().Replace('/', '\\').TrimEnd('\\');

            // if file version equals to "NA" then do not generate it.
            if (pattern.FileVersion.Trim().ToUpper() == "NA" ||
                 pattern.FileVersion.Trim().ToUpper() == "N/A")
            {
                patSetStatus.ValidTs = "NoVailidTs";
                LpatSetStatus.Add(pattern.PatternName, patSetStatus);
                noExistPatHash.Add(pattern.PatternName);
                return;
            }
            patSetStatus.ValidTs = "VailidTs";

            // Check if file exist
            bool patExisted = false;
            string path = patternFolder;
            if (string.IsNullOrEmpty(path))
            {
                path = Path.Combine(@"K:\", LocalSpecs.Project);

                CheckPatternInFileList(path, baseFilePath, ref patExisted, ref fileValue);
            }
            else
            {
                CheckPatternInFolder(pattern, patternFolder, patterns, ref patExisted, ref fileValue);
            }


            if (!patExisted)
            {
            }
            if (_dicPatPath != null)
            {
                string pat = Path.GetFileName(fileValue).ToUpper();
#pragma warning disable Ter402 // IGXL output path construction with intentional backslash conversion
                fileValue = @".\Pattern" + _dicPatPath[pat.Replace(_patExtName, "").Replace(".GZ", "")].Replace("/", "\\") + pat;
#pragma warning restore Ter402
            }

            else
            {
                // always put vm_vector name no matter subroutine existed or not. 2017/05/05 Osprey Team.
                if (!Regex.IsMatch(fileValue, ".gz:|" + _patExtName + ":", RegexOptions.IgnoreCase))
                {
                    fileValue = fileValue + ":" +
                                fileValue.Split(new[] { '\\', '/' }).Last().ToUpper().Replace(_patExtName, "").Replace(".GZ", "");
                }

                fileValue = fileValue.Replace(path, PatPath).Replace(@"\\", @"\").Replace("/", "\\");
            }
            fileValue = fileValue.Replace(path, PatPath).Replace(@"\\", @"\").Replace("/", "\\");

            patSetStatus.ContainSubr = "NonContainSubr";

            patSetStatus.Existed = !patExisted ? "NonExisted" : "Existed";

            LpatSetStatus.Add(pattern.PatternName, patSetStatus);

            // Change PatSet toupper and remove .gz
            var patSet = new PatSetRow
            {
                Burst = DefaultBurst.ToUpper(),
                File = fileValue.ToUpper().Replace(".GZ", "")
            };

            var patSetItem = new PatSet { PatSetName = pattern.PatternName.ToUpper() };

            if (patSetStatus.Existed == "NonExisted")
            {
                patSet.Comment = "NonExisted";
            }

            patSetItem.AddRow(patSet);

            if (patSetStatus.Existed == "Existed" || patSetStatus.Existed == "Skipped")
            {
                PatSetSheetAll.AddRow(patSetItem);
            }
            else
            {
                _patSetSheetAllnoPat.AddRow(patSetItem);
            }
        }

        private void CheckPatternInFileList(string path, string baseFilePath, ref bool patExisted, ref string fileValue)
        {
            if (_fileListinAllFolder.Count == 0)
            {
                //var files = new string[] { };

                //try
                //{
                //    allRelatedFile.ForEach(p => p = string.Format(@"{0}\{1}.PAT.gz", path, p));
                //    files = Directory.GetFiles(path, "*.PAT.*", SearchOption.AllDirectories);

                //}
                //catch (Exception e)
                //{
                //    Console.WriteLine(e.ToString());
                //}

                //foreach (var tmp in files)
                //{
                //    var fileName = tmp.Split('\\').Last().ToUpper();
                //    if (!_fileListinAllFolder.ContainsKey(fileName))
                //        _fileListinAllFolder.Add(fileName, tmp);
                //    else
                //    {

                //        var temp = compile[fileName.Split('.')[0]];

                //        var dupFileName = files.ToList().FindAll(p => Regex.IsMatch(p, fileName, RegexOptions.IgnoreCase));
                //        foreach (var checkFile in dupFileName.Where(checkFile => Regex.IsMatch(checkFile, temp.TpCategory, RegexOptions.IgnoreCase)))
                //        {
                //            _fileListinAllFolder[fileName] = checkFile;
                //            break;
                //        }
                //    }
                //}
                Dictionary<string, CompileItem> compilePat = _GetPatsetInfo();
                if (compilePat.Count != 0)
                {
                    foreach (KeyValuePair<string, CompileItem> compilepat in compilePat)
                    {
                        _fileListinAllFolder.Add(compilepat.Key + _patExtName + ".gz",
                            $"{path}\\{compilepat.Value.TpCategory}\\{compilepat.Key}{_patExtName}.gz");
                    }
                }
            }

            if (_fileListinAllFolder.ContainsKey(baseFilePath.ToUpper()))
            {
                patExisted = true;
                fileValue = _fileListinAllFolder[baseFilePath.ToUpper()];
            }
        }

        private void CheckPatternInFolder(PatternData pattern, string patternFolder, Dictionary<string, string> patterns, ref bool patExisted, ref string fileValue)
        {
            string patternName = Path.GetFileName(pattern.FileVersion).ToUpper().Replace(".ATP.GZ", _patExtName + ".GZ");
            patternName = patterns.ContainsKey(patternName.ToUpper()) ? patternName : Path.GetFileName(pattern.FileVersion).ToUpper().Replace(".ATP.GZ", _patExtName);

            if (patterns.ContainsKey(patternName.ToUpper()))
            {
                patterns.TryGetValue(patternName.ToUpper(), out string find);
                patExisted = true;

                _ = Path.GetFileName(patternFolder);
                fileValue = find.Replace(patternFolder, @".\" + "pattern");
            }

            //if (patterns.Exists(x => Path.GetFileName(x).Equals(patternName, StringComparison.CurrentCultureIgnoreCase)))
            //{
            //    var find = patterns.Find(x => Path.GetFileName(x).Equals(patternName, StringComparison.CurrentCultureIgnoreCase));
            //    patExisted = true;
            //    var patternFolderName = Path.GetFileName(patternFolder);
            //    fileValue = find.Replace(patternFolder, @".\" + "pattern");
            //}
        }

        private static string CreateFileValue(string fileVersion, out string baseFilePath)
        {
            // string patternPath = this.GetFileDir(org, typeSpec);
            string fileName = GetFileName(fileVersion);
            baseFilePath = fileName;
            return fileName;
        }

        /// <summary>
        /// Example 1: fileVersion sync_data/versions/1/CZ_MINA0_A_FULP_AN_AA02_FRQ_JTG_FRQ_ALLFV_SI_MEAS17_1_A0_1510261029.*.*
        /// Example 2: fileVersion CZ_MINA0_A_FULP_AN_AA02_FRQ_JTG_FRQ_ALLFV_SI_MEAS17_1_A0_1510261029.*.*
        /// Example 3: fileVersion NA  => ""
        /// </summary>
        /// <param name="fileVersion"></param>
        /// <returns></returns>
        public static string GetFileName(string fileVersion)
        {
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
                fileName += PatSuffix;
            }

            if (fileName.Trim().ToUpper() == "NA" || fileName.Trim().ToUpper() == "NA")
            {
                fileName = string.Empty;
            }

            return fileName;
        }

        public static PatSetSheet Convert2PatsetSheet(PatternSetSheet patsets)
        {
            var output = new PatSetSheet("PatSets_All");
            foreach (PatternSet patsetRow in patsets.PatternSets)
            {
                var patset = new PatSet { PatSetName = patsetRow.Name };

                foreach (PatternSetEntry entry in patsetRow.PatternEntries)
                {
                    patset.AddRow(new PatSetRow
                    {
                        Comment = entry.Comment,
                        File = entry.Pattern,
                        Burst = patsetRow.Burst,
                        Enable = entry.Enable,
                        StartLabel = entry.StartLabel,
                        StopLabel = entry.StopLabel,
                        TdGroup = entry.TDGroup,
                        TimeDomain = entry.TimeDomain
                    });
                }
                output.AddRow(patset);
            }
            return output;

        }

        private Dictionary<string, CompileItem> _GetPatsetInfo()
        {
            var result = new Dictionary<string, CompileItem>();
            string compilerFile = CopyCompilerResult(LocalSpecs.Project);
            //: string.Format(@"{0}\{1}_CompiledPat.csv", LocalSpecs.PatternPath, LocalSpecs.CurrentProject);
            //@"K:\{0}\Hard_IP\CSV_HardIP_Info\HardIP_AutoGen_Info_All.txt",


            if (File.Exists(compilerFile))
            {
                //Read Compile File
                var compileReader = new CompilePatReader();
                result = compileReader.ReadCompileFile(compilerFile);
            }
            return result;
        }

        private string CopyCompilerResult(string project)
        {
            string filename = $"{project}_CompiledPat.csv";
            string src = $@"\\STDP-SVN-C651\VPN\Log\CompiledPat\{filename}";
            if (!File.Exists(src))
            {
                //src = string.Format(@"\\STDP-SVN-m.tsmc\VPN\Log\{0}", filename);
                return "";
                //MessageBox.Show(
                //    "The main server is not available, takes from slave one! The pattern infomation might not be up to date!",
                //    "Warning");
            }
            string dlnPath = Path.Combine(Path.GetTempPath(), filename);
            File.Copy(src, dlnPath, true);
            return dlnPath;
        }

        private static void ModifyPatSetAll(PatSetSheet patsets)
        {
            foreach (PatSet patset in patsets.Rows)
            {
                int rowIndex = LocalSpecs.PatSetAll.Rows.FindIndex(p => p.PatSetName.Equals(patset.PatSetName, StringComparison.OrdinalIgnoreCase));

                if (rowIndex != -1)
                {
                    LocalSpecs.PatSetAll.Rows[rowIndex] = patset;
                }
                else
                {
                    LocalSpecs.PatSetAll.Rows.Add(patset);
                }
            }
        }
    }
    #endregion
}
