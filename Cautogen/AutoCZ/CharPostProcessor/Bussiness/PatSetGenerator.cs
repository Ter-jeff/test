using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class PatSetGenerator
    {
        private const string PatSetsAllCz = "PatSets_All_CZ";
        private const string PatternSubroutine = "Pattern_Subroutine";
        private const string PatSetAllSheet = "PatSets_All";
        private const string PatSetsCz = "PatSets_CZ";
        private static string _patternListcsvFile = "";
        private static string _charplanFile = "";
        private static string _charBaseProgram = "";

        private readonly string _patternFolder;
        private readonly Dictionary<string, SubrPatInfo> _hardIpInfoAllDict;
        private readonly bool _genCharNotUse;
        private readonly bool _genPatNotUse;
        private readonly bool _genPatSub;
        private readonly HashSet<string> _noExistPatHash = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _noUsePatHash = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public PatSetGenerator(InputParam inputParam)
        {
            _patternFolder = inputParam.PatternFolder;
            _hardIpInfoAllDict = inputParam.HardIpInfoAllDict;
            _genCharNotUse = inputParam.GenCharNotUse;
            _genPatNotUse = inputParam.GenPatNotUse;
            _genPatSub = inputParam.GenPatSub;
        }

        public void Generate(bool isExtraNeedAdd = true)
        {
            try
            {
                LocalSpecs.MessageWriter.WriteLine("Generating PatSetAll sheets...");
                _charplanFile = string.IsNullOrEmpty(LocalSpecs.InputParam.CharPlan)
                    ? Path.GetFileName(LocalSpecs.InputParam.CharFile.Remove(LocalSpecs.InputParam.CharFile.IndexOf("_Result_")) + ".xlsx")
                    : Path.GetFileName(LocalSpecs.InputParam.CharPlan);
                _patternListcsvFile = Path.GetFileName(LocalSpecs.InputParam.PatListFile);
                _charBaseProgram = Path.GetFileName(LocalSpecs.InputParam.ProgWorkBookPath);

                string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                    ? LocalSpecs.OutputFolder
                    : Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);

                string mainPatSetFile = LocalSpecs.ProgramUpdateOnly
                    ? Path.Combine(LocalSpecs.OutputFolder, ConstData.PatSetFolder, PatSetAllSheet + ".txt")
                    : Path.Combine(outputFolder, PatSetsAllCz + ".txt");

                //if the patsetall not from input of imfile, generate another patset by pattern list
                var allpatterns = LocalSpecs.CharPlanSheets.SelectMany(p => p.CharList)
                  .SelectMany(p => p.UsedPatterns).SelectMany(s => s.Split(',')).Distinct().ToList();
                string patExtName = ResolvePatExtensionName();
                List<string> patterns = ResolvePatternFiles(ref patExtName);
                Dictionary<string, string> patternsDict = BuildPatternDict(patterns);

                if (LocalSpecs.PatSetAll == null || LocalSpecs.PatSetAll.Rows.Count == 0)
                {
                    var patsetallgen = new PatSetAllGenerator(_genPatNotUse);
                    patsetallgen.GenerateFlow(mainPatSetFile, _patternFolder, patternsDict, allpatterns, _noUsePatHash, _noExistPatHash, patExtName);
                }

                #region get from current PatSetAll in test program
                PatSetSheet currentPatSetAll = LocalSpecs.TestProgram.PatSetSheets.FirstOrDefault(p => p.Name.Equals(PatSetAllSheet, StringComparison.OrdinalIgnoreCase));
                if (currentPatSetAll == null)
                {
                    return;
                }

                var currentPatterns = new HashSet<string>(currentPatSetAll.Rows.SelectMany(x => x.PatSetRows)
                                                            .Select(x => x.File)
                                                            .Select(x => Path.GetFileNameWithoutExtension(x.Split(':').First().ToUpper())));
                ModifyPatSetAll(currentPatSetAll, isExtraNeedAdd);
                LocalSpecs.PatSetAll.Rows.SelectMany(x => x.PatSetRows).Where(x => x.IsBackup).ToList().ForEach(x => _noExistPatHash.Add(x.PatternSet));
                #endregion

                if (LocalSpecs.PatSetAll != null)
                {
                    if (!LocalSpecs.ProgramUpdateOnly)
                    {
                        LocalSpecs.PatSetAll.Name = PatSetsAllCz;
                    }

                    LocalSpecs.PatSetAll.Write(mainPatSetFile, "2.3");
                    LocalSpecs.GenSheets.Add(LocalSpecs.PatSetAll);
                }

                #region patSetSubSheet
                if (!_genPatSub)
                {
                    return;
                }

                GeneratePatSetSub(patternsDict, currentPatterns, patExtName);
                #endregion
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.StackTrace);
                LogHelper.Error(ex.Message);
            }
        }

        private string ResolvePatExtensionName()
        {
            return Directory.Exists(Path.Combine(_patternFolder, "patx")) ? ".PATX" : ".PAT";
        }

        private List<string> ResolvePatternFiles(ref string patExtName)
        {
            var patterns = new List<string>(Directory.GetFiles(_patternFolder, $"*{patExtName}.gz", SearchOption.AllDirectories)).ToList();

            if (patterns.Count == 0)
            {
                patterns = new List<string>(Directory.GetFiles(_patternFolder, $"*{patExtName}", SearchOption.AllDirectories)).ToList();
            }

            if (patterns.Count == 0)
            {
                patExtName = patExtName == ".PAT" ? ".PATX" : patExtName;
                patterns = new List<string>(Directory.GetFiles(_patternFolder, $"*{patExtName}", SearchOption.AllDirectories)).ToList();
            }
            return patterns;
        }

        private static Dictionary<string, string> BuildPatternDict(List<string> patterns)
        {
            var patternsDict = new Dictionary<string, string>();

            foreach (string pattern in patterns)
            {
                if (!patternsDict.ContainsKey(Path.GetFileName(pattern).ToUpper()))
                {
                    patternsDict.Add(Path.GetFileName(pattern).ToUpper(), pattern);
                }
            }
            return patternsDict;
        }

        private void GeneratePatSetSub(Dictionary<string, string> patternsDict, HashSet<string> currentPatterns, string patExtName)
        {
            LocalSpecs.MessageWriter.WriteLine("Generating PatSetSub sheets...");
            PatSetSubSheet patSetSubSheet = LocalSpecs.TestProgram.PatSetSubSheets.FirstOrDefault() ?? new PatSetSubSheet(PatternSubroutine);
            //var patSetSubSheet = new PatSetSubSheet(PatternSubroutine);
            foreach (PatternData pattern in LocalSpecs.PatternDatas.Values)
            {
                if (pattern.PatternName.ToUpper().StartsWith("PP_COLA0_C_PLLP_BI_C0FC_BIR_JTG_59N_MEXXXX_SI_REP"))
                {
                }
                if (pattern.Use.ToUpper().Equals(PatSetAllGenerator.UsedType) ||
                _genPatNotUse ||
                LocalSpecs.PatternsInCharPlan.Contains(pattern.PatternName, StringComparer.InvariantCultureIgnoreCase))
                {
                    ProcessPatSetSubPattern(pattern, patternsDict, currentPatterns, patExtName, patSetSubSheet);
                }
            }
            string path = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.PatSetFolder);

            patSetSubSheet.Write(Path.Combine(path, patSetSubSheet.Name + ".txt"));
            LocalSpecs.GenSheets.Add(patSetSubSheet);
        }

        private void ProcessPatSetSubPattern(PatternData pattern, Dictionary<string, string> patternsDict, HashSet<string> currentPatterns, string patExtName, PatSetSubSheet patSetSubSheet)
        {
            string name = pattern.GetFileVersionNameOnly();
            string patName = name + patExtName + ".GZ";

            patternsDict.TryGetValue(patName.ToUpper(), out string find);
            //find = patterns.Find(x => Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase));
            if (find == null)
            {
                return;
            }

            string patsub = ResolvePatSub(name, pattern, find);
            if (string.IsNullOrEmpty(patsub))
            {
                patSetSubSheet.Rows.RemoveAll(x => x.PatternFileName.ToUpper().Contains(pattern.PatternName.ToUpper())
                                                            && currentPatterns.Contains(Path.GetFileNameWithoutExtension(x.PatternFileName.Split(':')[0]).ToUpper()));
                return;
            }

            //var patternFolderName = Path.GetFileName(patternFolder);
            string patternFolderName = "pattern";
            string fileValue = find.Replace(_patternFolder, @".\" + patternFolderName);
            fileValue = fileValue.ToUpper().Replace(".ATP.GZ", "").Replace(patExtName + ".GZ", "");

            //if (!currentPatterns.Exists(x => x.Equals(name, StringComparison.CurrentCultureIgnoreCase)))
            PatSetSubRow patrow =
                    patSetSubSheet.Rows.FirstOrDefault(
                        x => x.PatternFileName.ToUpper().Contains(pattern.PatternName.ToUpper()));
            if (patrow == null)
            {
                patrow = new PatSetSubRow { Comment = "New for CZ", PatternFileName = fileValue + ":" + patsub.ToUpper() };
                patSetSubSheet.Rows.Add(patrow);
            }
            else
            {
                if (patrow.PatternFileName != fileValue + patExtName + ":" + patsub.ToUpper())
                {
                    patrow.Comment = "Update version by CZ dashboard";
                    patrow.PatternFileName = fileValue + patExtName + ":" + patsub.ToUpper();
                }
            }
        }

        private string ResolvePatSub(string name, PatternData pattern, string find)
        {
            if (_hardIpInfoAllDict != null && _hardIpInfoAllDict.TryGetValue(name, out SubrPatInfo value))
            {
                return string.Join(",", value.Subroutine);
            }
            if (_hardIpInfoAllDict != null && _hardIpInfoAllDict.ContainsKey(pattern.PatternName.ToLower()))
            {
                return string.Join(",", _hardIpInfoAllDict[pattern.PatternName.ToLower()].Subroutine);
            }
            string atpContent = "";
            var patInforReader = new PatPatInforReader();
            if (!new PatInfoCmd().ConvertByArgs(find, ref atpContent, "-hdr -switches"))
            {
                return "";
            }
            string vmVectorName = patInforReader.VmVectorReader(atpContent.Split(new[] { '\n' }).ToList());
            if (vmVectorName.Split(',').Length == 2)
            {
                return vmVectorName.Split(',').Last();
            }
            return "";
        }

        public void GeneratePatSetsCz(List<CharPlanSheet> charPlanSheets)
        {
            var patSetsCz = new PatSetSheet("PatSets_CZ");
            foreach (CharPlanSheet planSheet in charPlanSheets)
            {
                if (planSheet.IsHardIp)
                {
                    continue;
                }

                bool enableMtd = planSheet.CharList.Any(p => !string.IsNullOrEmpty(p.Die));

                foreach (CharPlanItem planItem in planSheet.CharList)
                {
                    var patSetInit = new PatSet();
                    var patSetPayload = new PatSet();

                    bool hasInitNoUseExist = planItem.UsedInits.Any(x => _noUsePatHash.Contains(x) || _noExistPatHash.Contains(x));
                    hasInitNoUseExist |= !(planItem.Use || _genCharNotUse);

                    string initPatSetName = enableMtd ? $"{planItem.TestInstanceName}_{planItem.Die}_INIT" : planItem.TestInstanceName + "INIT";


                    patSetInit.PatSetName = initPatSetName;
                    planItem.UsedInits.ForEach(x => patSetInit.AddRow(new PatSetRow()
                    {
                        PatternSet = $"{planItem.TestInstanceName}{"INIT"}".ToUpper(),
                        Burst = "NO",
                        File = x.ToUpper(),
                        Comment = PrintNoUseExistInComment($"{planItem.SheetName}:Row{planItem.RowNum}", x, !(planItem.Use || _genCharNotUse)),
                        IsBackup = hasInitNoUseExist,
                    }));

                    string payloadPatSetName = enableMtd ? $"{planItem.TestInstanceName}_{planItem.Die}_PL" : planItem.TestInstanceName + "PL";


                    bool hasPlNoUseExist = planItem.UsedPayloads.Any(x => _noUsePatHash.Contains(x) || _noExistPatHash.Contains(x));
                    hasPlNoUseExist |= !(planItem.Use || _genCharNotUse);
                    patSetPayload.PatSetName = payloadPatSetName;
                    planItem.UsedPayloads.ForEach(x => patSetPayload.AddRow(new PatSetRow()
                    {
                        PatternSet = $"{planItem.TestInstanceName}{"INIT"}".ToUpper(),
                        Burst = "NO",
                        File = x.ToUpper(),
                        Comment = PrintNoUseExistInComment($"{planItem.SheetName}:Row{planItem.RowNum}", x, !(planItem.Use || _genCharNotUse)),
                        IsBackup = hasPlNoUseExist,
                    }));

                    if (patSetInit.PatSetRows.Any())
                    {
                        patSetsCz.AddRow(patSetInit);
                    }

                    if (patSetPayload.PatSetRows.Any())
                    {
                        patSetsCz.AddRow(patSetPayload);
                    }
                }
                string outPath = "";
                if (LocalSpecs.InputParam.GenTxtOnly)
                {
                    outPath = Path.Combine(LocalSpecs.OutputFolder, PatSetsCz + ".txt");
                }
                else
                {
                    outPath = Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder, PatSetsCz + ".txt");
                }

                patSetsCz.Write(outPath, "2.3");
                LocalSpecs.GenSheets.Add(patSetsCz);
            }
        }

        private string PrintNoUseExistInComment(string ori, string patternName, bool charNotUse)
        {
            string result = ori;
            if (_noExistPatHash.Contains(patternName))
            {
                result += ", NonExisted";
            }

            if (_noUsePatHash.Contains(patternName))
            {
                result += ", PatNotUse";
            }

            if (charNotUse)
            {
                result += ", CharNotUse";
            }

            return result;
        }

        public static void GenerateBak(string patternFolder, Dictionary<string, SubrPatInfo> hardIpInfoAllDict)
        {
            try
            {
                var usedPatterns = LocalSpecs.CharPlanSheets.SelectMany(p => p.CharList)
                .SelectMany(p => p.UsedPatterns).SelectMany(s => s.Split(',')).Select(x => x.ToUpper()).Distinct().ToList();

                var patterns = new List<string>(Directory.GetFiles(patternFolder, "*.gz", SearchOption.AllDirectories)).ToList();

                PatSetSheet patSetAll = LocalSpecs.TestProgram.PatSetSheets.FirstOrDefault(p => p.Name.Equals(PatSetAllSheet, StringComparison.OrdinalIgnoreCase));
                List<string> currentPatterns = patSetAll == null ?
                    null :
                    patSetAll.Rows.SelectMany(x => x.PatSetRows).Select(x => x.File).
                    Select(x => Path.GetFileNameWithoutExtension(x.Split(':').First())).ToList();

                var patSets = new List<PatSet>();
                var patSetSubRows = new List<PatSetSubRow>();
                foreach (string usedPattern in usedPatterns)
                {
                    if (LocalSpecs.PatternDatas.TryGetValue(usedPattern, out PatternData patternData))
                    {
                        string name = patternData.GetFileVersionNameOnly();
                        if (currentPatterns != null && currentPatterns.Exists(x => x.Equals(name, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            continue;
                        }

                        string patName = name + ".PAT.GZ";
                        if (patterns.Exists(x => Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            string find = patterns.Find(x => Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase));
                            string patternFolderName = Path.GetFileName(patternFolder);
                            string fileValue = find.Replace(patternFolder, @".\" + patternFolderName);
                            fileValue = fileValue.ToUpper().Replace(".ATP.GZ", "").Replace(".PAT.GZ", "");

                            var patSet = new PatSet { PatSetName = usedPattern };
                            var patSetRow = new PatSetRow
                            {
                                PatternSet = usedPattern,
                                Burst = "NO",
                                File = fileValue + ".PAT:" + name,
                                Comment = "New for CZ"
                            };
                            patSet.AddRow(patSetRow);
                            patSets.Add(patSet);

                            string patsub = "";
                            if (hardIpInfoAllDict != null && hardIpInfoAllDict.TryGetValue(name, out SubrPatInfo value))
                            {
                                patsub = string.Join(",", value.Subroutine);
                            }
                            else
                            {
                                string atpContent = "";
                                var patInforReader = new PatPatInforReader();
                                if (new PatInfoCmd().ConvertByArgs(find, ref atpContent, "-hdr -switches"))
                                {
                                    string vmVectorName =
                                        patInforReader.VmVectorReader(atpContent.Split(new[] { '\n' }).ToList());
                                    if (vmVectorName.Split(',').Length == 2)
                                    {
                                        patsub = vmVectorName.Split(',').Last();
                                    }
                                }
                            }
                            if (string.IsNullOrEmpty(patsub))
                            {
                                continue;
                            }

                            var patrow = new PatSetSubRow
                            {
                                Comment = "New for CZ",
                                PatternFileName = fileValue + ".PAT:" + patsub.ToUpper()
                            };
                            patSetSubRows.Add(patrow);
                        }
                    }
                }

                string patSetsAllCzFileName = Path.Combine(Path.Combine(LocalSpecs.OutputFolder, ConstData.PatSetFolder), PatSetsAllCz + ".txt");
                var patSetSheet = new PatSetSheet(PatSetsAllCz);
                patSetSheet.AddRows(patSets);
                if (patSetAll != null)
                {
                    foreach (PatSet row in patSetAll.Rows)
                    {
                        if (!patSetSheet.Rows.Exists(x => x.PatSetName.Equals(row.PatSetName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            patSetSheet.AddRow(row);
                        }
                    }
                    string patSetAllFileName = Path.Combine(Path.Combine(LocalSpecs.OutputFolder, ConstData.PatSetFolder), patSetAll.Name + ".txt");
                    patSetAll.Write(patSetAllFileName);
                }
                patSetSheet.Write(patSetsAllCzFileName);

                var subroutine = new PatSetSubSheet(PatternSubroutine);
                subroutine.AddRows(patSetSubRows);
                string patSetSubFileName = Path.Combine(Path.Combine(LocalSpecs.OutputFolder, ConstData.PatSetFolder), subroutine.Name + ".txt");
                if (LocalSpecs.TestProgram.PatSetSubSheets != null && LocalSpecs.TestProgram.PatSetSubSheets.Any())
                {
                    foreach (PatSetSubRow row in LocalSpecs.TestProgram.PatSetSubSheets.First().Rows)
                    {
                        if (!subroutine.Rows.Exists(x => x.PatternFileName.Equals(row.PatternFileName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            subroutine.Rows.Add(row);
                        }
                    }
                }
                subroutine.Write(patSetSubFileName);
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.StackTrace);
                LogHelper.Error(ex.Message);
            }
        }

        private static void ModifyPatSetAll(PatSetSheet patsetallTp, bool isExtraNeedAdd)
        {
            //var PatSetInfoChangeList = new 
            var commentList = patsetallTp.Rows.TakeWhile(item => !string.IsNullOrEmpty(item.PatSetRows[0].Comment)).Select(item => item.PatSetRows[0].Comment).ToList();

            var patternIssues = new List<PatSetChange>();

            foreach (PatSet patsetCsv in LocalSpecs.PatSetAll.Rows)
            {
                int rowIndex = patsetallTp.Rows.FindIndex(p => p.PatSetName.Equals(patsetCsv.PatSetName, StringComparison.OrdinalIgnoreCase));
                if (rowIndex != -1)
                {
                    PatSet temp = patsetallTp.Rows[rowIndex];
                    if (!temp.PatSetRows[0].File.Equals(patsetCsv.PatSetRows[0].File, StringComparison.OrdinalIgnoreCase))
                    {
                        var issueItem = new PatSetChange
                        {
                            Generic = temp.PatSetRows[0].PatternSet,
                            PatternInTp = temp.PatSetRows[0].File,
                            PatternInCsvServer = patsetCsv.PatSetRows[0].File
                        };
                        patternIssues.Add(issueItem);
                    }
                    patsetallTp.Rows[rowIndex] = patsetCsv;

                }
                else if (isExtraNeedAdd)
                {
                    patsetallTp.Rows.Add(patsetCsv);
                    var issueItem = new PatSetChange
                    {
                        Generic = patsetCsv.PatSetName,
                        PatternInCsvServer = patsetCsv.PatSetRows[0].File
                    };
                    patternIssues.Add(issueItem);
                }
            }

            commentList.Add("ver : Char_plan: " + _charplanFile);
            commentList.Add("ver : Char_dashboard: " + _patternListcsvFile);
            commentList.Add("ver : Char_base_program: " + _charBaseProgram);

            int appendCount = commentList.Count - patsetallTp.Rows.Count;
            if (appendCount > 0)
            {
                for (int i = 0; i < appendCount; i++)
                {
                    var newSubRow = new PatSetRow();
                    var newRow = new PatSet();
                    newRow.AddRow(newSubRow);
                    patsetallTp.AddRow(newRow);
                }
            }
            int appendIndex = 0;
            foreach (string appendComment in commentList)
            {
                patsetallTp.Rows[appendIndex].PatSetRows[0].Comment = appendComment;
                appendIndex++;
            }

            ExcelWorksheet workSheet = LocalSpecs.PostReportWriter.Workbook.Worksheets.Add("PatSetChangeSummary");
            int rowindex = 1;
            PrintCheckChangeHeader(workSheet, ref rowindex);
            foreach (PatSetChange issue in patternIssues.Where(p => !string.IsNullOrEmpty(p.PatternInTp)))
            {
                PrintCheckChangeInfo(workSheet, issue, ref rowindex);
            }

            PrintCheckEnd(workSheet, ref rowindex);

            PrintCheckExtraPatternHeader(workSheet, ref rowindex);
            foreach (PatSetChange issue in patternIssues.Where(p => string.IsNullOrEmpty(p.PatternInTp)))
            {
                PrintCheckExtraPatternInfo(workSheet, issue, ref rowindex);
            }
            PrintCheckEnd(workSheet, ref rowindex);
            try
            {
                LocalSpecs.PostReportWriter.Save();
            }
            catch (Exception)
            {

            }
            finally
            {
                LocalSpecs.PostReportWriter.Dispose();
            }
            LocalSpecs.PatSetAll = patsetallTp;
        }

        private static void PrintCheckChangeHeader(ExcelWorksheet sheet, ref int rowindex)
        {
            sheet.Cells[rowindex, 1].Value = "There are changes of pattern version list below";
            rowindex++;
            sheet.Cells[rowindex, 1].Value = "===============================================";
            rowindex++;
            sheet.Cells[rowindex, 1].Value = "Generic Pattern";
            sheet.Cells[rowindex, 2].Value = "Pattern version in test program";
            sheet.Cells[rowindex, 3].Value = "Pattern version from csv/server";
            rowindex++;
        }

        private static void PrintCheckEnd(ExcelWorksheet sheet, ref int rowindex)
        {
            sheet.Cells[rowindex, 1].Value = "===============================================";
            rowindex++;
        }

        private static void PrintCheckChangeInfo(ExcelWorksheet sheet, PatSetChange item, ref int rowindex)
        {
            sheet.Cells[rowindex, 1].Value = item.Generic;
            sheet.Cells[rowindex, 2].Value = item.PatternInTp;
            sheet.Cells[rowindex, 3].Value = item.PatternInCsvServer;
            rowindex++;
        }

        private static void PrintCheckExtraPatternHeader(ExcelWorksheet sheet, ref int rowindex)
        {
            sheet.Cells[rowindex, 1].Value = "There are added patterns list below";
            rowindex++;
            sheet.Cells[rowindex, 1].Value = "===============================================";
            rowindex++;
            sheet.Cells[rowindex, 1].Value = "Generic Pattern";
            sheet.Cells[rowindex, 2].Value = "Full pattern from csv/server";
            rowindex++;
        }

        private static void PrintCheckExtraPatternInfo(ExcelWorksheet sheet, PatSetChange item, ref int rowindex)
        {
            sheet.Cells[rowindex, 1].Value = item.Generic;
            sheet.Cells[rowindex, 2].Value = item.PatternInCsvServer;
            rowindex++;
        }
    }

    internal class PatSetChange
    {
        public string Generic;
        public string PatternInTp;
        public string PatternInCsvServer;
    }
}
