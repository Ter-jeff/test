using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

using Automation.Reader;
using Automation.Static;
using Automation.Utility.CollectPat;

using LogLib.Utility;

using OfficeOpenXml;

using ScghLib.Base;
using ScghLib.Reader;

using TestPlanLib.Basic;

namespace Automation.Utility.Pattern
{
    public class PatMbistInfoManager
    {
        private readonly string _patListCsv;
        private readonly List<BistProdFlowSheet> _bistFlowSheets;
        public string OutputPath;
        public string FileName;
        public List<string> LMbistInfo;
        public Dictionary<string, List<string>> UsagePattern = new Dictionary<string, List<string>>();
        public Dictionary<string, int> BlockMaxDepth = new Dictionary<string, int>();

        public ModeSel Mode;
        public int SocOffset { get; set; }
        public int CpuOffset { get; set; }
        public int GfxOffset { get; set; }
        public bool SocMore1Block { get; set; }
        public bool CpuMore1Block { get; set; }
        public bool GfxMore1Block { get; set; }
        public bool IgnoreDontCare { get; set; }
        public bool IsUfp { get; set; }
        public string MbistInfoCfgFile { get; set; }

        public void ReadMbistConfig()
        {
            var reader = new XmlSerializer(typeof(MbistInfoCfg));
            var cfgXml = new StreamReader(MbistInfoCfgFile);
            var overviewReader = (MbistInfoCfg)reader.Deserialize(cfgXml);
            cfgXml.Close();
            UpdateTheInfo(overviewReader);
        }

        private void UpdateTheInfo(MbistInfoCfg bistInfoCfg)
        {
            SocOffset = bistInfoCfg.SocOffset;
            CpuOffset = bistInfoCfg.CpuOffset;
            GfxOffset = bistInfoCfg.GfxOffset;
            SocMore1Block = bistInfoCfg.SocMoreThan1Block;
            CpuMore1Block = bistInfoCfg.CpuMoreThan1Block;
            GfxMore1Block = bistInfoCfg.GfxMoreThan1Block;
            Mode.Prod = bistInfoCfg.Pp;
            Mode.Debug = bistInfoCfg.Dd;
            Mode.Fa = bistInfoCfg.Fa;
            Mode.Char = bistInfoCfg.Cz;
            IgnoreDontCare = bistInfoCfg.IgnoreDontCare;
        }

        public PatMbistInfoManager(string patListCsv, List<string> prodFlowSheetList = null, bool isUfp = false)
        {
            _patListCsv = patListCsv;
            _bistFlowSheets = new List<BistProdFlowSheet>();
            IsUfp = isUfp;

            if (prodFlowSheetList != null)
            {
                foreach (string sheet in prodFlowSheetList)
                {
                    ExcelWorksheet worksheet = EpWorkbook.ScghWorkbook.Worksheets[sheet];
                    if (worksheet == null)
                    {
                        continue;
                    }

                    try
                    {
                        var reader = new BistProdFlowReader(new MbistSheet { SheetName = worksheet.Name });
                        BistProdFlowSheet result = reader.ReadSheet(worksheet);
                        _bistFlowSheets.Add(result);
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                }
            }

            OutputPath = "";
            Mode = new ModeSel();
            SocOffset = 0;
            CpuOffset = 0;
            GfxOffset = 0;
            SocMore1Block = false;
            CpuMore1Block = false;
            GfxMore1Block = false;
            IgnoreDontCare = false;
        }

        public bool GetMbistInfoFromServer(string hardipInfo, string bistInfo)
        {
            string patSetPath = LocalSpecs.PatternFolder;
            string hardipFolder = patSetPath + Path.DirectorySeparatorChar;
            Dictionary<string, SubrPatInfo> hardIpInfoAll = ResolveHardIpInfoAll(hardipInfo, hardipFolder);

            var fingerPrintList = new List<string>();
            if (!File.Exists(bistInfo))
            {
                return false;
            }

            var mbistReader = new MbistInfoReader(bistInfo, hardIpInfoAll, IsUfp);
            mbistReader.CollectItems();
            var patList = PatternListReader.GetPatternListDic(_patListCsv).Select(x => x.Value).ToList();
            LMbistInfo = new List<string>();
            if (_bistFlowSheets != null && _bistFlowSheets.Count > 0)
            {
                CollectFromProdFlowSheets(mbistReader, patList, fingerPrintList);
            }
            else
            {
                CollectFromPatternList(mbistReader, patList, fingerPrintList);
            }

            LocalSpecs.FingerPrintList = fingerPrintList;
            return true;
        }

        private Dictionary<string, SubrPatInfo> ResolveHardIpInfoAll(string hardipInfo, string hardipFolder)
        {
            var hardIpInfoAll = new Dictionary<string, SubrPatInfo>();
            if (string.IsNullOrEmpty(hardipInfo))
            {
                foreach (string hipInfo in Directory.GetFiles(hardipFolder, "HardIP_AutoGen_Info_All*"))
                {
                    hardIpInfoAll = ServerInfo.ReadHardIpInfoAll(hipInfo);
                }
            }
            else
            {
                if (File.Exists(hardipInfo))
                {
                    hardIpInfoAll = ServerInfo.ReadHardIpInfoAll(hardipInfo);
                }
                else
                {
                    if (Directory.Exists(hardipInfo))
                    {
                        foreach (string hipInfo in Directory.GetFiles(LocalSpecs.HardIpInfoFileName, "HardIP_AutoGen_Info_All*"))
                        {
                            hardIpInfoAll = ServerInfo.ReadHardIpInfoAll(hipInfo);
                        }
                    }
                }
            }

            return hardIpInfoAll;
        }

        private void CollectFromProdFlowSheets(MbistInfoReader mbistReader, List<PatternData> patList, List<string> fingerPrintList)
        {
            var patternsInScgh = new List<string>();
            foreach (BistProdFlowSheet prodFlowSheet in _bistFlowSheets)
            {
                patternsInScgh.AddRange(prodFlowSheet.Rows.Select(p => p.Pattern).Distinct().ToList());
            }
            foreach (string pattern in patternsInScgh)
            {
                PatternData targetItem = patList.Find(p => p.PatternName.Equals(pattern, StringComparison.OrdinalIgnoreCase));
                if (targetItem == null)
                {
                    continue;
                }

                string key = targetItem.FileVersion.Trim().ToUpper() == "N/A" || targetItem.FileVersion.Trim().ToUpper() == "NA"
                    ? "NA"
                    : targetItem.FileVersion.ToUpper().Split('/').Last().Replace(".ATP", "").Replace(".GZ", "").Replace(".PAT", "");
                if (mbistReader.LMbistInfo.TryGetValue(key, out List<string> value))
                {
                    List<string> filter = FilterByConfig(value);
                    if (filter.Any())
                    {
                        //Add pattern into FingerPrint list
                        fingerPrintList.Add(pattern);
                        LMbistInfo.AddRange(filter);
                        UsagePattern.Add(key, filter);
                    }
                }
            }
        }

        private void CollectFromPatternList(MbistInfoReader mbistReader, List<PatternData> patList, List<string> fingerPrintList)
        {
            foreach (PatternData patData in patList)
            {
                string key = patData.FileVersion.Trim().ToUpper() == "N/A" || patData.FileVersion.Trim().ToUpper() == "NA"
                    ? "NA"
                    : patData.FileVersion.ToUpper().Split('/').Last().Replace(".ATP", "").Replace(".GZ", "").Replace(".PAT", "");
                if (!Regex.IsMatch(key.ToUpper(), "DD|CZ|PP|FA"))
                {
                    continue;
                }

                if (IsFilteredOutByMode(key))
                {
                    continue;
                }

                if (!patData.Use.Equals("use", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (mbistReader.LMbistInfo.TryGetValue(key, out List<string> value))
                {
                    //Add pattern into FingerPrint list
                    List<string> filter = FilterByConfig(value);
                    if (filter.Any())
                    {
                        //Add pattern into FingerPrint list
                        fingerPrintList.Add(patData.PatternName);
                        LMbistInfo.AddRange(filter);
                        UsagePattern.Add(key, filter);
                    }
                }
            }
        }

        private bool IsFilteredOutByMode(string key)
        {
            if (!Mode.Prod && Regex.IsMatch(key, "PP"))
            {
                return true;
            }

            if (!Mode.Char && Regex.IsMatch(key, "CZ"))
            {
                return true;
            }

            if (!Mode.Debug && Regex.IsMatch(key, "DD"))
            {
                return true;
            }

            if (!Mode.Fa && Regex.IsMatch(key, "FA"))
            {
                return true;
            }

            return false;
        }

        internal List<string> FilterByConfig(List<string> lMbistInfo)
        {
            var list = new List<string>();
            string block = "";
            foreach (string mbistInfo in lMbistInfo)
            {
                string info = mbistInfo;
                string[] tmp = mbistInfo.Split('\t');
                if (IgnoreDontCare && tmp[4].ToUpper().Equals("X"))
                {
                    continue;
                }

                block = tmp[0].Split('_')[2];
                if (SocOffset != 0 || CpuOffset != 0 || GfxOffset != 0)
                {
                    info = ApplyCycleOffset(tmp, block);
                }
                if (IsFilteredByBlockPartCount(block, tmp))
                {
                    continue;
                }

                list.Add(info);
            }

            UpdateBlockMaxDepth(block, list.Count);
            return list;
        }

        private string ApplyCycleOffset(string[] tmp, string block)
        {
            string cycle = tmp[3];
            int.TryParse(cycle, out int iCycle);
            if (block.ToUpper() == "S" && SocOffset != 0)
            {
                iCycle -= SocOffset;
            }

            if (block.ToUpper() == "C" && CpuOffset != 0)
            {
                iCycle -= CpuOffset;
            }

            if (block.ToUpper() == "L" && GfxOffset != 0)
            {
                iCycle -= GfxOffset;
            }

            return tmp[0] + "\t" + tmp[1] + "\t" + tmp[2] + "\t" + iCycle + "\t" + tmp[4] + "\t" + tmp[5];
        }

        private bool IsFilteredByBlockPartCount(string block, string[] tmp)
        {
            if (block.ToUpper() == "S" && SocMore1Block && tmp[1].Split('_').Length <= 1)
            {
                return true;
            }

            if (block.ToUpper() == "C" && CpuMore1Block && tmp[1].Split('_').Length <= 1)
            {
                return true;
            }

            if (block.ToUpper() == "L" && GfxMore1Block && tmp[1].Split('_').Length <= 1)
            {
                return true;
            }

            return false;
        }

        private void UpdateBlockMaxDepth(string block, int listCount)
        {
            if (block != "")
            {
                if (!BlockMaxDepth.ContainsKey(block.ToUpper()))
                {
                    BlockMaxDepth.Add(block.ToUpper(), listCount);
                }
                else
                {
                    if (BlockMaxDepth[block.ToUpper()] < listCount)
                    {
                        BlockMaxDepth[block.ToUpper()] = listCount;
                    }
                }
            }
        }

        public void SaveMbistInfo(bool showMsg = false)
        {
            const int maxCount = 190000;
            if (LMbistInfo != null && LMbistInfo.Count > maxCount)
            {
                for (int i = 0; i < LMbistInfo.Count; i += maxCount)
                {
                    List<string> result = LMbistInfo.GetRange(i, Math.Min(maxCount, LMbistInfo.Count - i));
                    string newFileName = Path.GetFileNameWithoutExtension(FileName) + "_" + (i / maxCount) + Path.GetExtension(FileName);
                    WriteToFile(newFileName, result);

                    string newFileNameWithModuleName = Path.GetFileNameWithoutExtension(newFileName) + "_ModuleNameOnly" +
                                                       Path.GetExtension(FileName);
                    WriteToFile(newFileNameWithModuleName, result, true);

                }
            }
            else
            {
                WriteToFile(FileName, LMbistInfo);
                string newFileNameWithModuleName = Path.GetFileNameWithoutExtension(FileName) + "_ModuleNameOnly" +
                                                   Path.GetExtension(FileName);
                WriteToFile(newFileNameWithModuleName, LMbistInfo, true);
            }
        }

        public void SaveFingerPrintMaxDepth(string filename)
        {
            var sw = new StreamWriter(Path.Combine(OutputPath, filename));
            const string header = "Block\tMaxDepth";
            sw.WriteLine(header);
            foreach (KeyValuePair<string, int> item in BlockMaxDepth)
            {
                sw.WriteLine(item.Key + "\t" + item.Value);
            }
            sw.Close();
        }

        private void WriteToFile(string fileName, List<string> infoList, bool moduleNameOnly = false)
        {
            const string header = "Pattern\tBlock\tVector\tCycle\tCompare\tType";
            using (var sw = new StreamWriter(Path.Combine(OutputPath, fileName)))
            {
                sw.WriteLine(header);
                foreach (string mbistInfo in infoList)
                {
                    if (moduleNameOnly)
                    {
                        string info = mbistInfo.Split(':').Last();
                        sw.WriteLine(info);
                    }
                    else
                    {
                        sw.WriteLine(mbistInfo);
                    }
                }
            }
        }
    }
}
