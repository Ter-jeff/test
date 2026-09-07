using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using TestPlanLib;

namespace BinCutScriptLib.Reader
{
    public class BinCutDatalogConfigReader
    {
        private const string BcStartVbt = "*print: BinCut Config start*";
        private const string BcEndVbt = "*print: BinCut Config end*";

        public const string BcStart = "[INFO]  ----- BinCut Config start -----";
        private const string BcEnd = "[INFO]  ----- BinCut Config end -----";

        protected static void GetBinCutConfig(StreamReader streamReader)
        {
            var bcConfigDict = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            ReadBinCutConfig(streamReader, ref bcConfigDict);

            if (bcConfigDict.Count > 0)
            {
                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Mapping_TestJobName_to_BincutJobName), out string? value))
                {
                    BinCutConfig.MappingTestJobNameToBincutJobName = value;
                }

                BinCutConfig.VddbinningFailStopFlag = bcConfigDict.TryGetValue(nameof(EnumBcConfig.Flag_Vddbinning_Fail_Stop), out string? value1) ? value1 : "";
                BinCutConfig.VddbinningPowerBinningFailStopFlag = bcConfigDict.TryGetValue(nameof(EnumBcConfig.Flag_Vddbinning_Power_Binning_Fail_Stop), out string? value2) ? value2 : "";
                BinCutConfig.VddbinningIdsFailFlag = bcConfigDict.TryGetValue(nameof(EnumBcConfig.Flag_Vddbinning_IDS_fail), out string? value3) ? value3 : "";
                BinCutConfig.VddbinningInterpolationFailFlag = bcConfigDict.TryGetValue(nameof(EnumBcConfig.Flag_Vddbinning_Interpolation_fail), out string? value4) ? value4 : "";
                BinCutConfig.GradeSearchMethodSelected = bcConfigDict.TryGetValue(nameof(EnumBcConfig.GradeSearchMethod_Selected), out string? value5) ? value5 : "Conventional";
                BinCutConfig.IsDoAll = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.IGXL_RunOptions_DoAll));
                BinCutConfig.IsOverrideFailstop = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.IGXL_RunOptions_OverrideFailstop));
                BinCutConfig.IsEnablePowerBinningHarvest = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_Enable_PowerBinning_Harvest));
                BinCutConfig.FlagIdsDistributionEnable = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_IDS_Distribution_enable));
                BinCutConfig.IsCompareByProductValueOnly = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_Only_Check_PV_for_VoltageHeritage));
                BinCutConfig.FlagUseNewInterpolationMonotonicity = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_use_new_Interpolation_Monotonicity));
                BinCutConfig.FlagUseCofInstance = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_use_COFInstance));
                BinCutConfig.VddbinCofStepInheritance = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Vddbin_COF_StepInheritance));
                BinCutConfig.FlagCp1IdsNotUseHarvResult = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_CP1_IDS_NotUseHarvResult));
                BinCutConfig.FlagSkipPrintingSafeVoltage = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_Skip_Printing_Safe_Voltage));
                BinCutConfig.FlagSkipPrintingSelSrmDsscInfo = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_Skip_Printing_SelSrm_DSSC_Info));
                BinCutConfig.FlagCrossDomainCheckEnable = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_cross_domain_check_enable));
                BinCutConfig.FlagPowerBinningHarvestBinFieldName = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_PowerBinning_Harvest_Bin_FieldName));
                BinCutConfig.FlagVddbinHarvestBin4RunBin1Eqn1 = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_Vddbin_HarvestBin4_Run_Bin1EQN1));
                BinCutConfig.FlagNewSelsrmVthEnable = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_New_Selsrm_Vth_enable));
                BinCutConfig.FlagT0TxHotFormat = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_T0TX_HOT_format));
                BinCutConfig.FlagT0TxRoomFormat = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_T0TX_ROOM_format));
                BinCutConfig.FlagSyncUpDcvsOutputEnable = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Flag_SyncUp_DCVS_Output_enable_Printing));
                GetConfigHarvFlags(bcConfigDict);
                BinCutConfig.VddbinCofStepInheritanceNewLogic = GetConfigResult(bcConfigDict, nameof(EnumBcConfig.Vddbin_COF_StepInheritance_New_Logic));
                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Vdd_Binning_Def_appA_1), out string? value6))
                {
                    BinCutConfig.VddBinningDefAppA1 = value6;
                }

                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Vdd_Binning_Def_appA_2), out string? value7))
                {
                    BinCutConfig.VddBinningDefAppA2 = value7;
                }

                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Vdd_Binning_Def_appA_3), out string? value8))
                {
                    BinCutConfig.VddBinningDefAppA3 = value8;
                }

                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Non_Binning_Rail), out string? value9))
                {
                    BinCutConfig.NonBinningRail = value9;
                }

                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Power_Binning_Seq_Sheet), out string? value10))
                {
                    BinCutConfig.PowerBinningSeqSheet = value10;
                }

                if (bcConfigDict.TryGetValue(nameof(EnumBcConfig.Is_BinCutJob_for_StepSearch), out string? value11))
                {
                    BinCutConfig.IsBinCutJobForStepSearch = value11.EqualsIgnoreCase("true");
                }
                else
                {
                    BinCutConfig.IsBinCutJobForStepSearch = null;
                }
            }
            else
            {
                BinCutConfig.IsDoAll = false;
            }
        }

        public static void ReadBinCutConfig(string dataLogFile, ref Job job)
        {
            using (StreamReader sr = GetStreamReader(dataLogFile))
            {
                GetBinCutConfig(sr);
            }

            if (!string.IsNullOrEmpty(BinCutConfig.MappingTestJobNameToBincutJobName))
            {
                job.JobType = new Job(BinCutConfig.MappingTestJobNameToBincutJobName).JobType;
            }
        }

        public static void ReadBinCutConfigCs(string dataLogFile, ref Job job, string tempFolder)
        {
            var configDict = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            using (StreamReader sr = GetStreamReader(dataLogFile))
            {
                configDict = ReadBinCutConfigCsharp(sr);
            }

            if (configDict.Count > 0)
            {
                #region Job
                if (configDict.TryGetValue(nameof(EnumBcConfig.Mapping_TestJobName_to_BincutJobName), out string? value))
                {
                    BinCutConfig.MappingTestJobNameToBincutJobName = value;
                }
                else if (configDict.ContainsKey("CP1"))
                {
                    BinCutConfig.MappingTestJobNameToBincutJobName = "CP1";
                }
                #endregion

                if (configDict.TryGetValue(nameof(EnumBcConfig.bincut_eqn_appA), out string? value12))
                {
                    BinCutConfig.BincutEqnAppA = value12;
                }

                if (configDict.TryGetValue(nameof(EnumBcConfig.bincut_ate_condition_non), out string? value13))
                {
                    BinCutConfig.BincutAteConditionNon = value13;
                }

                if (configDict.TryGetValue(nameof(EnumBcConfig.Power_Binning_Seq_Sheet), out string? value14))
                {
                    BinCutConfig.PowerBinningSeqSheet = value14;
                }

                BinCutConfig.DebugBinCutCofE1 = GetConfigResult(configDict, nameof(EnumBcConfig.Debug_BinCutCOF_E1));
                BinCutConfig.DebugBinCutCofStored = GetConfigResult(configDict, nameof(EnumBcConfig.Debug_BinCutCOF_Stored));
                //Csharp datalog no use it, Keep it for script procegression it define default = "Conventional"
                BinCutConfig.GradeSearchMethodSelected = "Conventional";

                // BinCutConfig.start_eqn = GetConfigResult(configDict, EnumBcConfig.start_eqn.ToString()); //ori is Flag_IDS_Distribution_enable

                BinCutConfig.FlagSkipPrintingSafeVoltage = GetConfigResult(configDict, nameof(EnumBcConfig.Flag_Skip_Printing_Safe_Voltage));
                BinCutConfig.FlagSkipPrintingSelSrmDsscInfo = GetConfigResult(configDict, nameof(EnumBcConfig.Flag_Skip_Printing_SelSrm_DSSC_Info));
                BinCutConfig.FlagCrossDomainCheckEnable = GetConfigResult(configDict, nameof(EnumBcConfig.Flag_cross_domain_check_enable));
                BinCutConfig.FlagUseNewInterpolationMonotonicity = GetConfigResult(configDict, nameof(EnumBcConfig.Flag_use_new_Interpolation_Monotonicity));
                BinCutConfig.FlagUseCofInstance = GetConfigResult(configDict, nameof(EnumBcConfig.Flag_use_COFInstance));
                //Default true

                BinCutConfig.FlagNewSelsrmVthEnable = GetConfigResultForDefaultTrue();
                //Default true
                BinCutConfig.IsEnablePowerBinningHarvest = GetConfigResultForDefaultTrue();
                //Default true
                BinCutConfig.IsCompareByProductValueOnly = GetConfigResultForDefaultTrue();

                if (configDict.TryGetValue(nameof(EnumBcConfig.Is_BinCutJob_for_StepSearch), out string? value15))
                {
                    BinCutConfig.IsBinCutJobForStepSearch = value15.EqualsIgnoreCase("true");
                }
                else
                {
                    BinCutConfig.IsBinCutJobForStepSearch = null;
                }

                GetConfigHarvFlagsCsharp(configDict);

            }
            else
            {
                BinCutConfig.IsDoAll = false;
            }

            if (!string.IsNullOrEmpty(BinCutConfig.MappingTestJobNameToBincutJobName))
            {
                job.JobType = new Job(BinCutConfig.MappingTestJobNameToBincutJobName).JobType;
            }

            using StreamReader sr2 = GetStreamReader(dataLogFile);
            BinCutConfig.StartEqn = GetStartEqnInfo(sr2);
        }

        private static void ProcessEnableWords(string dataLogFile)
        {
            BinCutConfig.ActiveEnableWords = ReadEnableWords(dataLogFile);

            //VB
            if (BinCutConfig.ActiveEnableWords.Contains("DebugPrintFlag"))
            {
                BinCutConfig.DebugPrintFlag = true;
            }

            //Csharp
            if (BinCutConfig.ActiveEnableWords.Contains("LogLevel_Debug"))
            {
                BinCutConfig.EnableWordLogLevelDebug = true;
            }
            if (BinCutConfig.ActiveEnableWords.Contains("LogLevel_NPI"))
            {
                BinCutConfig.EnableWordLogLevelNpi = true;
            }
            if (BinCutConfig.ActiveEnableWords.Contains("LogLevel_MP"))
            {
                BinCutConfig.EnableWordLogLevelMp = true;
            }
        }

        public static void ReadBinCutDatalogConfig(string dataLogFile, string tempFolder, ref Job job, string programName, bool csFlag)
        {
            if (!csFlag)
            {
                ReadBinCutConfig(dataLogFile, ref job);
                ProcessEnableWords(dataLogFile);
                BinCutConfig.MultiFstpEnable = ReadMultiFstpEnable(BinCutData.TestInstanceSheet!);
                CheckFiles(tempFolder);
                BinCutData.Job = job;
            }
            else
            {
                ReadBinCutConfigCs(dataLogFile, ref job, tempFolder);
                ProcessEnableWords(dataLogFile);
                //BinCutConfig.IsDoAll = GetConfigResultForDoall(dataLogFile);
                BinCutConfig.MultiFstpEnable = ReadMultiFstpEnable(BinCutData.TestInstanceSheet!);
                CheckFilesCs(tempFolder);
                BinCutData.Job = job;
            }

            SetBinCutT0TxFlags(programName, job);
        }

        public static void SetBinCutT0TxFlags(string programName, Job job)
        {
            // Set default true for safe voltage compare.
            //BinCutConfig.Flag_T0TX_ROOM_format = true;
            bool isT0Tx = programName.Contains("T0TX", StringComparison.OrdinalIgnoreCase);
            bool isRoomTemp = programName.Contains("25C", StringComparison.OrdinalIgnoreCase);
            bool isHotTemp = programName.Contains("85C", StringComparison.OrdinalIgnoreCase) || programName.Contains("105C", StringComparison.OrdinalIgnoreCase);
            bool isCp1OrFt1 = job.Equals("CP1") || job.Equals("FT1");

            if (isT0Tx && (isRoomTemp || isCp1OrFt1))
            {
                BinCutConfig.FlagT0TxRoomFormat = true;
                BinCutConfig.FlagT0TxHotFormat = false;
            }
            else if (isT0Tx && isHotTemp)
            {
                BinCutConfig.FlagT0TxHotFormat = true;
                BinCutConfig.FlagT0TxRoomFormat = false;
            }
        }

        public static void CheckFiles(string tempFolder)
        {
            CheckFile(tempFolder, BinCutConfig.NonBinningRail);

            CheckFile(tempFolder, BinCutConfig.VddBinningDefAppA1);

            CheckFile(tempFolder, BinCutConfig.VddBinningDefAppA2);

            CheckFile(tempFolder, BinCutConfig.VddBinningDefAppA3);
        }

        public static void CheckFilesCs(string tempFolder)
        {
            CheckFile(tempFolder, BinCutConfig.BincutAteConditionNon);

            CheckFile(tempFolder, BinCutConfig.BincutEqnAppA);
        }

        private static void CheckFile(string tempFolder, string fileName)
        {
            if (!string.IsNullOrEmpty(fileName))
            {
                string file = fileName + ".txt";
                string path = Path.Combine(tempFolder, file);
                if (!File.Exists(path))
                {
                    BinCutController.Controller.RichTextBoxAppend($"Can not find file {path} !!!", Color.Red);
                }
            }
        }

        public static void ReadBinCutConfig(StreamReader streamReader, ref Dictionary<string, string> binCutConfig)
        {
            bool bStart = false;
            int cnt = 0;
            while (!streamReader.EndOfStream)
            {
                string? line = streamReader.ReadLine();
                //prevent imcompelete datalog. start/stop isn't pair
                cnt++;
                if (cnt > 2000000)
                {
                    break;
                }

                if (line != null)
                {
                    string data = line.Trim();
                    if (data.EqualsIgnoreCase(BcStartVbt))
                    {
                        bStart = true;
                        continue;
                    }

                    if (bStart && data.EqualsIgnoreCase(BcEndVbt))
                    {
                        break;
                    }

                    if (line.Length != 0 && bStart)
                    {
                        if (Reg.RegexBincitVbtConfigValue.IsMatch(data))
                        {
                            string[] ary = line.StartsWith("site:") ? data.Split(':') : data.Split('=');
                            string name = line.StartsWith("site:") ? "HarvFlags " + ary[1].Split(',').First() : ary[0].Trim();

                            if (!binCutConfig.ContainsKey(name))
                            {
                                binCutConfig.Add(name, ary[1].Trim());
                            }
                        }
                    }
                }
            }
        }

        public static Dictionary<string, string> ReadBinCutConfigCsharp(StreamReader streamReader)
        {
            var config = new Dictionary<string, string>();
            var bcConfigRows = new List<string>();
            var siteFlags = new Dictionary<string, List<string>>();
            bool bStartOri = false;
            bool bStart = false;

            int cnt = 0;
            while (!streamReader.EndOfStream)
            {
                string? line = streamReader.ReadLine();
                //prevent imcompelete datalog. start/stop isn't pair
                cnt++;
                if (cnt > 2000000)
                {
                    break;
                }

                if (line != null)
                {
                    string data = line.Trim();

                    #region VBT Parsing Part
                    //If library update, it need to delete
                    if (data.EqualsIgnoreCase(BcStartVbt))
                    {
                        bStartOri = true;
                        continue;
                    }

                    if (bStartOri && data.EqualsIgnoreCase(BcEndVbt))
                    {
                        break;
                    }

                    if (line.Length != 0 && bStartOri)
                    {
                        if (Reg.RegexBincitVbtConfigValue.IsMatch(data))
                        {
                            string[] ary = line.StartsWith("site:") ? data.Split(':') : data.Split('=');
                            string name = line.StartsWith("site:") ? "HarvFlags " + ary[1].Split(',').First() : ary[0].Trim();

                            if (!config.ContainsKey(name))
                            {
                                config.Add(name, ary[1].Trim());
                            }
                        }
                    }
                    #endregion

                    if (data.EqualsIgnoreCase(BcStart))
                    {
                        bStart = true;
                        continue;
                    }
                    if (bStart && data.EqualsIgnoreCase(BcEnd))
                    {
                        break;
                    }

                    if (bStart)
                    {
                        HandleBStart(bcConfigRows, siteFlags, data);
                    }
                }
            }

            PostAction(config, bcConfigRows, siteFlags);

            return config;
        }

        private static void PostAction(Dictionary<string, string> config, List<string> bcConfigRows, Dictionary<string, List<string>> siteFlags)
        {
            foreach (KeyValuePair<string, List<string>> site in siteFlags)
            {
                bcConfigRows.Add($"{site.Key},{string.Join(",", site.Value)}");
            }

            foreach (string line in bcConfigRows)
            {
                string[] parts = line.Split('=');
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();

                    config.TryAdd(key, value);
                }
                else if (line.Contains("HarvFlags "))
                {
                    string[] spl = line.Split(',');
                    string key = spl[0];
                    string value = line.Replace(spl[0] + ",", "");

                    config.TryAdd(key, value);
                }
            }
        }

        private static void HandleBStart(List<string> bcConfigRows, Dictionary<string, List<string>> siteFlags, string data)
        {
            Match bincitConfigMatch = Reg.RegexBincitConfigValue.Match(data);
            Match harvFlagsMatch = Reg.RegexHarvFlags.Match(data);

            if (bincitConfigMatch.Success)
            {
                string key = bincitConfigMatch.Groups[1].Value.Trim();
                string value = bincitConfigMatch.Groups[2].Value.Trim();
                bcConfigRows.Add($"{key}={value}");
            }
            else if (harvFlagsMatch.Success)
            {
                string site = harvFlagsMatch.Groups[1].Value;
                string flagName = harvFlagsMatch.Groups[2].Value.Replace("'", "");
                string flagValue = harvFlagsMatch.Groups[3].Value.Trim();

                flagValue = flagValue.EqualsIgnoreCase("True") ? "T" : "F";

                string siteKey = $"HarvFlags {site}";

                if (!siteFlags.TryGetValue(siteKey, out List<string>? value))
                {
                    value = [];
                    siteFlags[siteKey] = value;
                }

                value.Add($"{flagName}={flagValue}");

            }
        }

        public static List<string> ReadEnableWords(string dataLogFile)
        {
            const string enableWords = "Active EnableWords";
            var enableWordsList = new List<string>();
            int cnt = 0;
            using (StreamReader sr = GetStreamReader(dataLogFile))
            {
                while (!sr.EndOfStream)
                {
                    string? line = sr.ReadLine();
                    cnt++;
                    if (cnt > 1000000)
                    {
                        break;
                    }

                    if (line != null && line.Contains(enableWords))
                    {
                        string enableWordsString = line.Replace("'", "").Split([':'], StringSplitOptions.RemoveEmptyEntries).Last();
                        enableWordsList = [.. enableWordsString.Split(['|'], StringSplitOptions.RemoveEmptyEntries)];
                        break;
                    }
                }
            }
            return enableWordsList;
        }

        public static bool ReadMultiFstpEnable(InstanceSheet instanceSheet)
        {
            string[] argArray = instanceSheet.Rows.First().ArgList.Split(',');
            int argIndex = Array.IndexOf(argArray, "MultiFSTP_Enable");

            return argIndex != -1 && instanceSheet.Rows.Exists(x => x.Args[argIndex] == "TRUE");
        }

        public static HarvestResult GetHarvestResult(string data)
        {
            string[] ary = data.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries);
            var result = new HarvestResult();
            if (ary[0].Contains(':'))
            {
                _ = int.TryParse(ary[0].Split(':').Last(), out int value);
                result.Site = value;
            }
            var harvestDic = new Dictionary<string, string>();
            foreach (string str in ary)
            {
                if (str.Contains('='))
                {
                    string[] harveStatus = str.Split('=');
                    if (harveStatus.Length != 0 && !harvestDic.ContainsKey(harveStatus.First()))
                    {
                        harvestDic.Add(harveStatus.First(), harveStatus.Last());
                    }
                }
            }
            if (harvestDic.Count != 0)
            {
                result.HarvesFlags = harvestDic;
            }

            return result;
        }

        private static StreamReader GetStreamReader(string dataLogFile)
        {
            string extension = Path.GetExtension(dataLogFile);
            return extension != null && extension.EqualsIgnoreCase(".txt") ?
                new StreamReader(File.OpenRead(dataLogFile)) :
                new StreamReader(new GZipStream(File.OpenRead(dataLogFile), CompressionMode.Decompress));
        }

        private static bool GetConfigResult(Dictionary<string, string> bcConfigDict, string key)
        {
            bool result = false;
            if (bcConfigDict.TryGetValue(key, out string? value))
            {
                result = value.EqualsIgnoreCase("true");
            }

            return result;
        }

        private static bool GetConfigResultForDefaultTrue()
        {
            bool result = true;
            return result;
        }

        private static void GetConfigHarvFlags(Dictionary<string, string> bcConfigDict)
        {
            // Match strict key formats containing "HarvFlags"
            ProcessConfigDictionary(bcConfigDict, key => key.Contains("HarvFlags"));
        }

        private static void GetConfigHarvFlagsCsharp(Dictionary<string, string> bcConfigDict)
        {
            // Match expanded C# formats containing "HarvFlags" or "harvesting"
            ProcessConfigDictionary(bcConfigDict, key => key.Contains("HarvFlags") || key.Contains("harvesting"));
        }

        private static void ProcessConfigDictionary(Dictionary<string, string> bcConfigDict, Func<string, bool> keyPredicate)
        {
            if (bcConfigDict == null)
            {
                return;
            }

            foreach (KeyValuePair<string, string> config in bcConfigDict)
            {
                if (config.Key != null && keyPredicate(config.Key))
                {
                    ParseAndRegisterFlags(config.Key, config.Value);
                }
            }
        }

        private static void ParseAndRegisterFlags(string key, string rawValue)
        {
            if (string.IsNullOrEmpty(rawValue))
            {
                return;
            }

            var flags = new List<string>();
            string[] flagList = rawValue.Split(',');

            foreach (string flag in flagList)
            {
                if (!string.IsNullOrEmpty(flag) && flag.Contains('='))
                {
                    string flagName = flag.Split('=').First();
                    if (!flags.Contains(flagName))
                    {
                        flags.Add(flagName);
                    }
                }
            }

            // Isolate the dictionary lookup target identifier
            string dictKey = key.Split(' ').Last();
            if (!string.IsNullOrEmpty(dictKey) && !BinCutConfig.HarvFlags.ContainsKey(dictKey))
            {
                BinCutConfig.HarvFlags.Add(dictKey, flags);
            }
        }

        private static bool GetStartEqnInfo(StreamReader streamReader)
        {
            var startEqnDict = new Dictionary<string, string>();
            var startEqnRows = new List<string>();
            bool bStart = false;

            int cnt = 0;
            while (!streamReader.EndOfStream)
            {
                string? line = streamReader.ReadLine();
                cnt++;
                if (cnt > 2000000)
                {
                    break;
                }

                if (line != null)
                {
                    string data = line.Trim();

                    if (data.EqualsIgnoreCase(BcStart))
                    {
                        bStart = true;
                        continue;
                    }
                    if (bStart && data.EqualsIgnoreCase(BcEnd))
                    {
                        break;
                    }

                    if (bStart)
                    {
                        Match startEqnMatch = Reg.RegexStartEqn.Match(data);

                        if (startEqnMatch.Success)
                        {
                            string key = startEqnMatch.Groups[1].Value.Trim();
                            string value = startEqnMatch.Groups[2].Value.Trim();
                            startEqnRows.Add($"{key}={value}");
                        }
                    }
                }
            }
            foreach (string row in startEqnRows)
            {
                string[] parts = row.Split('=');
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();

                    startEqnDict.TryAdd(key, value);
                }
            }
            return startEqnDict.Count > 0;
        }
    }
}
