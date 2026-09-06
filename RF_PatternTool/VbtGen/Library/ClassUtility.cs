using System.Text.RegularExpressions;

using Automation.Static;
using Automation.Utility.CollectPat;

using EfuseCheckCmdLib.IgxlLogLib;

using LogLib.Utility;

using TestPlanLib.Static;
using TestPlanLib.Xml;

using DataTable = System.Data.DataTable;

namespace Automation.Library
{
    public class FileList
    {
        public List<string> ItemListMap = new List<string>();
        public List<string> PathList = new List<string>();
        public List<string> Filelist = new List<string>();
        public string FilterStr = ".*";
        public void ProcessDirectory(string targetDirectory, FileList listPtr, bool dealFile = false)
        {
            if (listPtr != null)
            {
                listPtr.ItemListMap.Add(targetDirectory);
            }
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
            {
                if (dealFile)
                {
                    ProcessFile(fileName, listPtr);
                }
            }

            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory, listPtr, dealFile);
            }
        }

        public void ProcessFile(string path, FileList listPtr)
        {
            Console.WriteLine(@"Processed file '{0}'.", path);
            if (listPtr != null)
            {
                if (Regex.IsMatch(path, FilterStr, RegexOptions.IgnoreCase))
                {
                    listPtr.ItemListMap.Add(path);
                    listPtr.Filelist.Add(path);
                }
            }
        }
    }

    public class ClassUtility
    {
        public static string ScghScanRegStr()
        {
            string regstr = "";
            if (!string.IsNullOrEmpty(NeededSheets.ScanScgh))
            {
                regstr += NeededSheets.ScanScgh + "|";
            }
            if (!string.IsNullOrEmpty(NeededSheets.ScanScghCpu))
            {
                regstr += NeededSheets.ScanScghCpu + "|";
            }
            if (!string.IsNullOrEmpty(NeededSheets.ScanScghGpu))
            {
                regstr += NeededSheets.ScanScghGpu + "|";
            }
            if (!string.IsNullOrEmpty(NeededSheets.ScanScghSoc))
            {
                regstr += NeededSheets.ScanScghSoc + "|";
            }
            regstr = regstr.Substring(0, regstr.Length - 1);
            return regstr;
        }

        public static Dictionary<string, string> DifferentialPair(List<string> pinList)
        {
            DiffPairConfig config = XmlService<DiffPairConfig>.LoadXml(Directory.GetCurrentDirectory() + "/Config/DiffPairConfig.xml");
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            foreach (DiffItem pinPair in config.DiffPairPins)
            {
                //Add differential pairs from config which defined using Pin name
                string posPin = pinList.Find(p => p.Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase));
                string negPin = pinList.Find(p => p.Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase));
                if (posPin != null && negPin != null && !pairs.ContainsKey(posPin))
                {
                    pairs.Add(posPin, negPin);
                }
            }

            for (int i = 0; i < pinList.Count; i++)
            {
                for (int j = i + 1; j < pinList.Count; j++)
                {
                    if (pinList[i].Length == pinList[j].Length)
                    {
                        GetSamePartInDiffPairs(pinList[i], pinList[j], out string nStr, out string pStr);
                        bool flag = false;
                        foreach (DiffItem rule in config.DiffPairRules)
                        {
                            if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.ContainsKey(pinList[j]))
                                {
                                    pairs.Add(pinList[j], pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                            else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.ContainsKey(pinList[i]))
                                {
                                    pairs.Add(pinList[i], pinList[j]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                        }
                        if (flag)
                        {
                            break;
                        }
                    }
                }
            }

            //Store Neg and Pos pin again
            int count = pairs.Count;
            for (int i = 0; i < count; i++)
            {
                pairs.Add(pairs.ElementAt(i).Value, pairs.ElementAt(i).Key);
            }
            return pairs;
        }

        public static List<string> GroupDiffPairs(List<string> oriPinList)
        {
            var pinList = oriPinList.ToList();
            DiffPairConfig config = XmlService<DiffPairConfig>.LoadXml(Directory.GetCurrentDirectory() + "/Config/DiffPairConfig.xml");
            List<string> pairs = new List<string>();

            foreach (DiffItem pinPair in config.DiffPairPins)
            {
                //Add differential pairs from config which defined using Pin name
                string posPin = pinList.Find(p => p.Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase));
                string negPin = pinList.Find(p => p.Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase));
                if (posPin != null && negPin != null && !pairs.Contains(posPin + "::" + negPin))
                {
                    pairs.Add(posPin + "::" + negPin);
                    pinList.Remove(posPin);
                    pinList.Remove(negPin);
                }
            }

            for (int i = 0; i < pinList.Count; i++)
            {
                for (int j = i + 1; j < pinList.Count; j++)
                {
                    if (pinList[i].Length == pinList[j].Length)
                    {
                        GetSamePartInDiffPairs(pinList[i], pinList[j], out string nStr, out string pStr);
                        bool flag = false;
                        foreach (DiffItem rule in config.DiffPairRules)
                        {
                            if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.Contains(pinList[j] + "::" + pinList[i]))
                                {
                                    pairs.Add(pinList[j] + "::" + pinList[i]);
                                    pinList.Remove(pinList[j]);
                                    pinList.Remove(pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                            else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.Contains(pinList[i] + "::" + pinList[j]))
                                {
                                    pairs.Add(pinList[i] + "::" + pinList[j]);
                                    pinList.Remove(pinList[j]);
                                    pinList.Remove(pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                        }
                        if (flag)
                        {
                            break;
                        }
                    }
                }
            }

            pairs.AddRange(pinList);
            return pairs;
        }

        public static bool DiffPinPosAndNeg(string diffPins, out string pos, out string neg, out string groupName)
        {
            pos = "";
            neg = "";
            groupName = "";
            bool isDiff = false;
            if (!diffPins.Contains("::"))
            {
                return false;
            }

            string[] pair = diffPins.Split(new string[] { "::" }, StringSplitOptions.None);
            string diffPairConfig = string.IsNullOrEmpty(LocalSpecs.SettingFolder) ? Directory.GetCurrentDirectory() + "/Config/DiffPairConfig.xml" : LocalSpecs.SettingFolder + @"\config\DiffPairConfig.xml";
            DiffPairConfig config = XmlService<DiffPairConfig>.LoadXml(diffPairConfig);
            for (int i = 0; i < pair.Length; i++)
            {
                pair[i] = pair[i].Trim();
            }

            #region Diff group name by rule-1 in DiffPairPins
            foreach (DiffItem pinPair in config.DiffPairPins)
            {
                if (pair[0].Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase)
                    && pair[1].Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[0];
                    neg = pair[1];
                    groupName = pos;
                    isDiff = true;
                }
                if (pair[1].Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase)
                        && pair[0].Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[1];
                    neg = pair[0];
                    groupName = pos;
                    isDiff = true;
                }
            }
            #endregion

            #region Diff group name by rule-2 get common string and compare diff. string

            GetSamePartInDiffPairs(ref groupName, pair[0], pair[1], out string nStr, out string pStr);

            foreach (DiffItem rule in config.DiffPairRules)
            {
                if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) &&
                    nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                        StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[0];
                    neg = pair[1];
                    isDiff = true;
                }
                else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) &&
                         pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                             StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[1];
                    neg = pair[0];
                    isDiff = true;
                }
            }
            #endregion

            #region Diff group name by rule-3
            //EX:ADDR_M2P_DQ_N::ADDR_M2P_DQ_P
            foreach (DiffItem rule in config.DiffPairRules)
            {
                if (Regex.IsMatch(pair[0], rule.Pos, RegexOptions.IgnoreCase) &&
                    pair[1].Equals(Regex.Replace(pair[0], rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                        StringComparison.OrdinalIgnoreCase))
                {
                    string newGroupName = Regex.Replace(pair[0], rule.Pos, "");
                    groupName = newGroupName.Length < groupName.Length ? newGroupName : groupName;
                    pos = pair[0];
                    neg = pair[1];
                    isDiff = true;
                }
                else if (Regex.IsMatch(pair[1], rule.Pos, RegexOptions.IgnoreCase) &&
                         pair[0].Equals(Regex.Replace(pair[1], rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                             StringComparison.OrdinalIgnoreCase))
                {
                    string newGroupName = Regex.Replace(pair[1], rule.Pos, "");
                    groupName = newGroupName.Length < groupName.Length ? newGroupName : groupName;
                    pos = pair[1];
                    neg = pair[0];
                    isDiff = true;
                }
            }
            #endregion

            if (isDiff)
            {
                if (!string.IsNullOrEmpty(groupName))
                {
                    groupName += groupName.EndsWith("_", StringComparison.OrdinalIgnoreCase) ? "Diff" : "_Diff";
                }
            }
            else
            {
                groupName = "";
            }

            return isDiff;
        }

        private static void GetSamePartInDiffPairs(ref string groupName, string nPinName, string pPinName, out string nStr, out string pStr)
        {
            //EX:ADDR_M2P_DQ_N0::ADDR_M2P_DQ_P0
            pStr = "";
            nStr = "";
            if (nPinName.Length == pPinName.Length)
            {
                for (int i = 0; i < nPinName.Length; i++)
                {
                    if (nPinName[i] == pPinName[i])
                    {
                        groupName += nPinName[i];
                    }
                    else
                    {
                        pStr += nPinName[i];
                        nStr += pPinName[i];
                    }
                }
            }
        }

        private static void GetSamePartInDiffPairs(string nPinName, string pPinName, out string nStr, out string pStr)
        {
            //EX:ADDR_M2P_DQ_N0::ADDR_M2P_DQ_P0
            pStr = "";
            nStr = "";
            if (nPinName.Length == pPinName.Length)
            {
                for (int i = 0; i < nPinName.Length; i++)
                {
                    if (nPinName[i] != pPinName[i])
                    {
                        nStr += nPinName[i];
                        pStr += pPinName[i];
                    }
                }
            }
        }
    }

    public class MbistInfoReader
    {
        private string _compileFile;
        private Dictionary<string, SubrPatInfo> _hardIpInfoAll;
        public Dictionary<string, List<string>> LMbistInfo;
        private bool _isUFP;
        public MbistInfoReader(string compileFile, Dictionary<string, SubrPatInfo> hardIpInfoAll, bool isUFP = false)
        {
            _compileFile = compileFile;
            _hardIpInfoAll = hardIpInfoAll;
            LMbistInfo = new Dictionary<string, List<string>>();
            _isUFP = isUFP;
            //CollectItems();
        }

        public void CollectItems()
        {
            if (File.Exists(_compileFile))
            {
                _ = new List<string>();
                List<string> fileDataList = ReadFileData();
                //get header index
                DateHeaderIndex headerIndex = HandleFileHeader(fileDataList[0]);
                //handle 5000 lines each time to avoid ContextSwitchDeadlock exception
                for (int i = 1; i < fileDataList.Count; i = i + 5000)
                {
                    int count = i + 5000 < fileDataList.Count ? 5000 : fileDataList.Count - i;
                    HandleFileData(headerIndex, fileDataList, i, count);
                }
            }
            else
            {
                //ErrorMessage.Show(
                //    @"Sorry, the project doesn't support Mbist Info, please check with IT team for details.", @"Warning",
                //    MessageBoxButton.OK);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private List<string> ReadFileData()
        {
            var lineList = new List<string>();
            var fileInfo = new FileInfo(_compileFile);
            string tmpFile = Path.Combine(Path.GetTempPath(), fileInfo.Name);
            File.Copy(_compileFile, tmpFile, true); // force overwrite 20160712
            var reader = new StreamReader(File.OpenRead(tmpFile));
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                lineList.Add(line);
            }
            reader.Close();
            File.Delete(tmpFile);
            return lineList;
        }

        private static DateHeaderIndex HandleFileHeader(string line)
        {
            DateHeaderIndex headerIndex = new DateHeaderIndex();

            string[] values = line.Split(',');
            if (Regex.IsMatch(line, ".*Pattern.*Block.*")) //the 1st line
            {
                for (int item = 0; item < values.Length; item++)
                {
                    if (values[item] == "Pattern")
                    {
                        headerIndex.Idx_pattern = item;
                    }
                    if (values[item] == "Block")
                    {
                        headerIndex.Idx_block = item;
                    }
                    if (values[item] == "Vector")
                    {
                        headerIndex.Idx_vector = item;
                    }
                    if (values[item] == "Cycle")
                    {
                        headerIndex.Idx_cycle = item;
                    }
                    if (values[item] == "Compare")
                    {
                        headerIndex.Idx_compare = item;
                    }
                    if (values[item] == "Type")
                    {
                        headerIndex.Idx_type = item;
                    }
                }
            }
            else
            {
                headerIndex = null;
            }

            return headerIndex;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lineList"></param>
        /// <param name="startRow"></param>
        /// <param name="count"></param>
        private void HandleFileData(DateHeaderIndex headerIndex, List<string> lineList, int startRow, int count)
        {
            if (headerIndex == null)
            {
                return;
            }

            int idx_pattern = headerIndex.Idx_pattern;
            int idx_block = headerIndex.Idx_block;
            int idx_vector = headerIndex.Idx_vector;
            int idx_cycle = headerIndex.Idx_cycle;
            int idx_compare = headerIndex.Idx_compare;
            int idx_type = headerIndex.Idx_type;

            var rgBlockCapture = new Regex(@"(?<Pick4>\w{1,4})\[(?<Num>\d+)\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);


            for (int i = startRow; i < startRow + count; i++)
            {
                string line = lineList[i];
                string[] values = line.Split(',');
                //string block = "";
                List<string> blocks = new List<string>();


                if (values[idx_pattern].Length > 0) //key
                {
                    string[] tmp = values[idx_block].Split(' ');

                    string currentBlock = "";

                    foreach (string subBlock in tmp)
                    {

                        Match m = rgBlockCapture.Match(subBlock);
                        if (m.Success)
                        {
                            string pick4Block = m.Groups["Pick4"].Value;
                            string num = m.Groups["Num"].Value;
                            if (currentBlock != pick4Block)
                            {
                                currentBlock = pick4Block;
                                blocks.Add(pick4Block + num);
                            }
                            else
                            {
                                blocks.Add(num);
                            }
                        }


                    }
                    string block = string.Join("_", blocks);

                    string lVmVector = "";
                    if (_hardIpInfoAll.ContainsKey(values[idx_pattern]))
                    {
                        lVmVector = _hardIpInfoAll[values[idx_pattern]].VmVector;
                    }

                    if (string.IsNullOrEmpty(lVmVector))
                    {
                        lVmVector = values[idx_pattern];
                    }

                    string fileType = (_isUFP) ? ".PATX:" : ".PAT:";
                    string lineTab = values[idx_pattern] + fileType + lVmVector + "\t" + block + "\t" + values[idx_vector] + "\t" +
                                  values[idx_cycle] + "\t" + values[idx_compare] + "\t" + values[idx_type];
                    if (!LMbistInfo.ContainsKey(values[idx_pattern].ToUpper()))
                    {
                        LMbistInfo.Add(values[idx_pattern].ToUpper(), new List<string>());
                    }

                    LMbistInfo[values[idx_pattern].ToUpper()].Add(lineTab);
                }

            }
        }
    }

    public class DateHeaderIndex
    {
        public int Idx_pattern;
        public int Idx_block;
        public int Idx_vector;
        public int Idx_cycle;
        public int Idx_compare;
        public int Idx_type;
    }

    public class StepEvaluator
    {
        //static Dictionary<string, int> dicUnit;
        public static string ErrMSg;
        //private static string _inStr;
        public static int Evaluate(string inStr)
        {
            ErrMSg = "";
            //_inStr = inStr;
            string[] tmp = inStr.Replace(" ", "").Split(',').ToArray();
            string start = Replace(Replace(tmp[0], @"-"), @"+");
            string stop = Replace(Replace(tmp[1], @"-"), @"+");
            if (!Regex.IsMatch(start, "^-"))
            {
                start = "+" + start;
            }

            if (!Regex.IsMatch(stop, "^-"))
            {
                stop = "+" + stop;
            }

            List<string> lstep = new List<string> { tmp[2] };
            List<string> lstart = Regex.Split(start, @",").ToList();
            List<string> lstop = Regex.Split(stop, @",").ToList();
            var lcheck = lstart.Select(a => a).ToList();
            foreach (string item in lcheck)
            {
                if (lstop.Contains(item))
                {
                    lstart.Remove(item);
                    lstop.Remove(item);
                }
            }
            //TODO                
            //`es
            double sweepRange = EvaluateExpression(GetCalStr(lstart, lstop));
            double stepSize = EvaluateExpression(CalStr(lstep));
            return Convert.ToInt32(sweepRange / stepSize);
        }
        private static double EvaluateExpression(string eqn)
        {
            var dt = new DataTable();
            try
            {
                object result = dt.Compute(eqn, string.Empty);
                return Convert.ToDouble(result);
            }
            catch (Exception ex)
            {
                ErrMSg = "Error! Can't do datatable EvaluateExpression! >> " + eqn;
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                return 0;
            }
        }
        private static string GetCalStr(List<string> l1, List<string> l2)
        {
            string evalStrSt = CalStr(l1);
            string evalStrSp = CalStr(l2);
            return evalStrSp + "-(" + evalStrSt + ")";
        }
        private static string CalStr(List<string> l1)
        {
            string evalStr = "";
            foreach (string item in l1)
            {
                string tmp = Regex.Replace(item, "V|A|Hz|OHM|S", "", RegexOptions.IgnoreCase);
                Match mc = Regex.Match(tmp, @"\d+(?<unit>[a-zA-Z])", RegexOptions.IgnoreCase);
                if (mc.Success)
                {
                    string scale = mc.Groups["unit"].ToString();
                    evalStr = evalStr + tmp.Replace(scale, "*" + Math.Pow(10, UnitUtit.Instance.DicUnit[mc.Groups["unit"].ToString()]).ToString());
                }
                else
                {
                    evalStr = evalStr + tmp;
                }
            }
            return evalStr;
        }
        private static string Replace(string inStr, string rpls)
        {
            inStr = Regex.Replace(inStr, "\\" + rpls, "," + rpls);
            return inStr;
        }

    }

}
