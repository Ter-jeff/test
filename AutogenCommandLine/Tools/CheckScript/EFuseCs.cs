using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;

using EfuseCheckCmdLib.Base;
using EfuseCheckCmdLib.CFGTable;
using EfuseCheckCmdLib.Datalog;
using EfuseCheckCmdLib.Output;

using OfficeOpenXml;

using TestPlanLib.Efuse;

namespace AutogenCommandLine.Tools.CheckScript
{
    public partial class EfuseAlgorithmCheckCs
    {
        private static string _inDir;
        private static string _bitDef;
        private static string _cfg;
        private static string _outputDir;

        private static string _datalog;
        private static string _bdfPath;
        private static string _cfgPath;
        private static string _outputFile;
        public static readonly Regex RegDram = DramRegex();
        public static readonly Regex RegFuseCheck = FuseCheckRegex();

        [GeneratedRegex("DRAM", RegexOptions.IgnoreCase)]
        private static partial Regex DramRegex();

        [GeneratedRegex("FuseCheck", RegexOptions.IgnoreCase)]
        private static partial Regex FuseCheckRegex();

        public EfuseAlgorithmCheckCs(string inDir, string bitDef, string cfg, string outputDir)
        {
            _inDir = inDir;
            _bitDef = bitDef;
            _cfg = cfg;
            _outputDir = outputDir;
        }

        public static void WorkFlow()
        {
            ErrorReportManager.ClearErrors();

            _datalog = _inDir.Trim('"');
            _bdfPath = _bitDef.Trim('"');
            _cfgPath = string.IsNullOrEmpty(_cfg) ? "" : _cfg.Trim('"');
            _outputFile = Path.Combine(_outputDir.Trim('"'), "EFuseReport");

            List<string> files = [.. Directory.GetFiles(_datalog, "*.txt", SearchOption.TopDirectoryOnly)];
            files.AddRange([.. Directory.GetFiles(_datalog, "*.gz", SearchOption.TopDirectoryOnly)]);
            if (files.Count == 0)
            {
                Console.WriteLine($"Error: Log file not found at {_datalog}");
                return;
            }

            if (!Directory.Exists(_outputFile))
            {
                Directory.CreateDirectory(_outputFile);
            }

            Console.WriteLine("Parsing Efuse_BitDef_Table in path : {0}", _bdfPath);
            var efuseBitDef = new LoaderEfuseBitDef(_bdfPath, "");
            efuseBitDef.Parse(new EfuseScriptConfig());
            Console.WriteLine("Parsing Efuse_BitDef_Table Done.");

            CfgTableReader.Reset();
            string[] cfgFiles = _cfgPath.Split([','], StringSplitOptions.RemoveEmptyEntries);
            foreach (string cfg in cfgFiles)
            {
                if (File.Exists(cfg))
                {
                    Console.WriteLine("Parsing Config_Table in path : {0}", cfg);
                    CfgTableReader.WorkFlow(cfg);
                    Console.WriteLine("Parsing Config_Table Done.");
                }
            }

            #region DRAM Type Check
            string directoryPath = Path.GetDirectoryName(_bdfPath);
            if (directoryPath != null)
            {
                var efuseDram = new EfuseDramTable(Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly).Where(file => file.EndsWith(".xlsx") || file.EndsWith(".txt")).ToArray()
                    .FirstOrDefault(file => RegDram.IsMatch(Path.GetFileName(file))), "");
                if (efuseDram.InPath != null)
                {
                    Console.WriteLine("Parsing DRAM Type Table in path : {0}", efuseDram.InPath);
                    efuseDram.Parse();
                    Console.WriteLine("Parsing DRAM Type Table Done.");
                }
                #endregion

                #region FuseCheck
                var fuseCheck = new FuseCheckTable(Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly).Where(file => file.EndsWith(".xlsx") || file.EndsWith(".txt")).ToArray()
                        .FirstOrDefault(file => RegFuseCheck.IsMatch(Path.GetFileName(file))), "");
                if (fuseCheck.InPath != null)
                {
                    Console.WriteLine("Parsing FuseCheck Table in path : {0}", fuseCheck.InPath);
                    fuseCheck.Parse();
                    Console.WriteLine("Parsing FuseCheck Table Done.");
                }
                #endregion

                foreach (string file in files)
                {
                    Console.WriteLine("Parsing datalog in path : {0}", file);
                    OneTouchDown touchDownHandler = new OneTouchDown(file);
                    Dictionary<string, int> diceXyInfo = [];
                    int xyNum = 0;
                    List<List<string>> segmentData = touchDownHandler.SplitDataFile();
                    foreach (List<string> segment in segmentData)
                    {
                        List<DiceInfo> allDices = new DatalogReader(segment).Read(efuseDram);
                        if (allDices.Count > 0)
                        {
                            string fileName = "";
                            foreach (DiceInfo dice in allDices)
                            {
                                if (diceXyInfo.ContainsKey(dice.XCoor + "_" + dice.YCoor))
                                {
                                    Console.WriteLine("Duplicate X Y values : {0}, bypass this loop", dice.XCoor + "_" + dice.YCoor);
                                    continue;
                                }

                                diceXyInfo.Add(dice.XCoor + "_" + dice.YCoor, xyNum);
                                xyNum++;

                                if (!string.IsNullOrEmpty(fileName))
                                {
                                    fileName += $"__X{dice.XCoor}_Y{dice.YCoor}";
                                }
                                else
                                {
                                    fileName += $"X{dice.XCoor}_Y{dice.YCoor}";
                                }
                            }

                            if (!string.IsNullOrEmpty(fileName))
                            {
                                string outputFile = Path.Combine(_outputFile, $"report_{fileName}.xlsx");
                                WriteExcel.WriteReport(allDices, efuseBitDef, efuseDram, fuseCheck, _cfgPath, outputFile);
                                Console.WriteLine("Generate report under path : {0}", outputFile);
                                using var excelPackage = new ExcelPackage(new FileInfo(outputFile));
                                ErrorReportManager.GenErrorReport(excelPackage, null, "ErrorReport");
                                excelPackage.Save();
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Efuse check done.");
        }
    }
}
