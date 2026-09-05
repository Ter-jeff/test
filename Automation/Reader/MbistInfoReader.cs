using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.Utility.CollectPat;

using LogLib.Utility;

namespace Automation.Reader
{
    public class MbistInfoReader
    {
        private readonly string _compileFile;
        private readonly Dictionary<string, SubrPatInfo> _hardIpInfoAll;
        public Dictionary<string, List<string>> LMbistInfo;
        private readonly bool _isUfp;
        public MbistInfoReader(string compileFile, Dictionary<string, SubrPatInfo> hardIpInfoAll, bool isUfp = false)
        {
            _compileFile = compileFile;
            _hardIpInfoAll = hardIpInfoAll;
            LMbistInfo = new Dictionary<string, List<string>>();
            _isUfp = isUfp;
            //CollectItems();
        }

        public void CollectItems()
        {
            if (File.Exists(_compileFile))
            {
                List<string> list = ReadFileData();
                //get header index
                DateHeaderIndex headerIndex = HandleFileHeader(list[0]);
                //handle 5000 lines each time to avoid ContextSwitchDeadlock exception
                for (int i = 1; i < list.Count; i += 5000)
                {
                    int count = i + 5000 < list.Count ? 5000 : list.Count - i;
                    HandleFileData(headerIndex, list, i, count);
                }
            }
            else
            {
                ErrorMessageBox.Show("Sorry, the project doesn't support Mbist Info, please check with IT team for details.", "Warning");
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

        private DateHeaderIndex HandleFileHeader(string line)
        {
            var headerIndex = new DateHeaderIndex();

            string[] values = line.Split(',');
            if (Regex.IsMatch(line, ".*Pattern.*Block.*")) //the 1st line
            {
                for (int item = 0; item < values.Length; item++)
                {
                    if (values[item] == "Pattern")
                    {
                        headerIndex.IdxPattern = item;
                    }

                    if (values[item] == "Block")
                    {
                        headerIndex.IdxBlock = item;
                    }

                    if (values[item] == "Vector")
                    {
                        headerIndex.IdxVector = item;
                    }

                    if (values[item] == "Cycle")
                    {
                        headerIndex.IdxCycle = item;
                    }

                    if (values[item] == "Compare")
                    {
                        headerIndex.IdxCompare = item;
                    }

                    if (values[item] == "Type")
                    {
                        headerIndex.IdxType = item;
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

            int idxPattern = headerIndex.IdxPattern;
            int idxBlock = headerIndex.IdxBlock;
            int idxVector = headerIndex.IdxVector;
            int idxCycle = headerIndex.IdxCycle;
            int idxCompare = headerIndex.IdxCompare;
            int idxType = headerIndex.IdxType;

            var rgBlockCapture = new Regex(@"(?<Pick4>\w{1,4})\[(?<Num>\d+)\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);


            for (int i = startRow; i < startRow + count; i++)
            {
                string line = lineList[i];
                string[] values = line.Split(',');
                var blocks = new List<string>();


                if (values[idxPattern].Length > 0) //key
                {
                    string[] tmp = values[idxBlock].Split(' ');

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
                    if (_hardIpInfoAll.ContainsKey(values[idxPattern]))
                    {
                        lVmVector = _hardIpInfoAll[values[idxPattern]].VmVector;
                    }

                    if (string.IsNullOrEmpty(lVmVector))
                    {
                        lVmVector = values[idxPattern];
                    }

                    string fileType = _isUfp ? ".PATX:" : ".PAT:";
                    string lineTab = values[idxPattern] + fileType + lVmVector + "\t" + block + "\t" + values[idxVector] + "\t" +
                                     values[idxCycle] + "\t" + values[idxCompare] + "\t" + values[idxType];
                    if (!LMbistInfo.ContainsKey(values[idxPattern].ToUpper()))
                    {
                        LMbistInfo.Add(values[idxPattern].ToUpper(), new List<string>());
                    }

                    LMbistInfo[values[idxPattern].ToUpper()].Add(lineTab);
                }

            }
        }
    }
}
