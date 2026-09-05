using System.Text.RegularExpressions;

using RF_PatternTool.PatternGen;

namespace RF_PatternTool.PatternStruct
{
    public class MatchLoopPattern
    {
        public List<string> RawDataList = new List<string>();
        private readonly double _freq;

        public int PatBit = 32;

        private static readonly string _jtagTdiSrcStr = "((JTAG_TDI):DigSrc = Send)";
        private static readonly string _jtagTdoCapStr = "((JTAG_TDO):DigCap = Store)";

        public MatchLoopPattern(double jtagFreq)
        {
            _freq = jtagFreq * 1e6;
        }
        //Read Raw data template
        public void ReadTemplate()
        {
            List<string> lines = Template.TemplateSet.MloopTemp.Replace("\r", "").Split('\n').ToList();

            foreach (string line in lines)
            {
                if (line.Contains(">"))
                {
                    RawDataList.Add(line.Split('/')[0]);
                }
                else
                {
                    RawDataList.Add(line);
                }
            }
        }
        public List<string> Write(BenchLogItem log, Register register, VectorInfo info, StreamWriter swlog)
        {
            List<string> result = new List<string>();
            string message = $"//OP:{log.Operation}. Address:{log.Address}. Field:{log.RegFieldName.Split('.').Last()}.";
            try
            {
                bool isRepeat = false;
                string sourceAddressStr = ConvertAddressLSBFirst(register.Address, 32) + ConvertAddressLSBFirst(register.Address, 32);// Address

                char[] dataList = ("".PadLeft(register.Data.Length, 'X')).ToCharArray();

                string binValue = Convert.ToString((long)(log.FieldVal.StartsWith("0x") ?
                        Convert.ToUInt64(log.FieldVal.Replace("0x", ""), 16) :
                        Convert.ToUInt64(log.FieldVal)), 2);
                binValue = binValue.PadLeft(PatBit, '0').Replace("0", "L").Replace("1", "H");// change to compare mode
                //reverse
                char[] revArr = binValue.ToCharArray();
                Array.Reverse(revArr);

                int shiftbit = 0;
                int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
                bool isNeedShift = PatBit == 64 && addrQuot % 2 != 0;
                if (isNeedShift)
                {
                    shiftbit = 32;
                }

                int index = 0;
                for (int i = log.LSB; i <= log.MSB; i++)
                {
                    dataList[i + shiftbit] = revArr[index];
                    index++;
                }

                int patAddIndex = -1;
                int patDataIndex = -1;
                //Assume 32 bits of data + 32 bits of Address
                //-> Assume 32 bits of Address + (32 or 64) bits of data + 32 bits of Address + (32 or 64) bits of data
                int tmp_rem = 0;
                string matchloopAddressMsg = $"MatchLoopAddr:{log.Address}";
                string matchloopDataMsg = $"MatchLoopData:{log.FieldVal}";
                for (int i = 0; i < RawDataList.Count; i++)
                {
                    string data = RawDataList[i];
                    if (data.Contains(_jtagTdiSrcStr))
                    {
                        patAddIndex++;
                        tmp_rem = Math.DivRem(patAddIndex, 32, out patAddIndex);

                        i++; //Skip source syntax, use fixed value
                        data = RawDataList[i].Replace(" D ", " " + sourceAddressStr.Substring(0, 1) + " ");
                        string addrMsg = AppendLogInfo(patAddIndex, matchloopAddressMsg, 32);

                        AddVectorRow(result, data, addrMsg, patAddIndex, info);

                        if (PatternGenBusiness.IsAddComment)
                        {
                            result.Add(string.Format("// address_csr[{0}]", patAddIndex)); //address_csr[dataindex-32]
                        }

                        sourceAddressStr = sourceAddressStr.Substring(1, sourceAddressStr.Length - 1); //remove LSB data
                    }
                    else if (data.Contains(_jtagTdoCapStr))
                    {
                        patDataIndex++;
                        tmp_rem = Math.DivRem(patDataIndex, PatBit, out patDataIndex);

                        string dataMsg = AppendLogInfo(patDataIndex, matchloopDataMsg, PatBit);
                        data = RawDataList[++i];
                        data = data.Replace(" V ", " " + dataList[patDataIndex] + " ");
                        if (dataList[patDataIndex] != 'L' && dataList[patDataIndex] != 'H')
                        {
                            data = data.Replace("mask", "");
                        }

                        AddVectorRow(result, data, dataMsg, patDataIndex, info);
                    }
                    else
                    {
                        WriteOtherRow(result, data, i, log, info, message, ref isRepeat);
                    }
                }
                info.Clear();
            }
            catch (Exception)
            {
                ;
            }

            return result;
        }

        private static void AddVectorRow(List<string> result, string data, string logMsg, int recordIndex, VectorInfo info)
        {
            string patrow = data;
            if (!string.IsNullOrEmpty(logMsg))
            {
                patrow += $"// {logMsg}";
            }

            if (recordIndex % 2 == 1)
            {
                patrow += $"// [{recordIndex}]";
            }

            if (PatternGenBusiness.IsAddComment)
            {
                patrow += info.GetVectorInfo();
                info.Update();
            }
            result.Add(patrow);
        }

        private void WriteOtherRow(List<string> result, string data, int rawIndex, BenchLogItem log, VectorInfo info, string message, ref bool isRepeat)
        {
            data = ExpandWaitRows(result, data, log, info);
            if (data.Contains("repeat"))
            {
                isRepeat = true;
                string regRepC = @"repeat\s+(?<count>\d+)";
                info.WaitCount = int.Parse(Regex.Match(data, regRepC, RegexOptions.IgnoreCase).Groups["count"].Value);
            }

            if (data.Contains(">"))
            {
                if (isRepeat && string.IsNullOrEmpty(info.LastVector))
                {
                    info.LastVector = data;
                }

                if (rawIndex == RawDataList.Count - 1)
                {
                    data = data + " " + message;
                }

                string patrow = data;
                if (PatternGenBusiness.IsAddComment)
                {
                    patrow += info.GetVectorInfo();
                    info.Update();
                }
                result.Add(patrow);
            }
            else
            {
                if (data.StartsWith("//"))
                {
                    string regData = @"data\[(?<num>\d+)\]";
                    if (data.ToLower().Contains("data") && Regex.IsMatch(data, regData, RegexOptions.IgnoreCase))
                    {
                        result.Add("// ");
                    }
                }
                result.Add(data);

            }

        }

        private string ExpandWaitRows(List<string> result, string data, BenchLogItem log, VectorInfo info)
        {
            if (data.Contains("Wait"))
            {
                double defV = Math.Ceiling(153 + _freq * (log.Wait));
                if (defV > 65536)
                {
                    int tmpQ;
                    for (int j = 0; j < Math.DivRem((int)defV, 65536, out tmpQ); j++)
                    {
                        string regRepC = @"repeat\s+(?<count>\d+)";
                        string tmpWait = data.Replace("Wait", "65536");
                        result.Add(string.Format(tmpWait));
                        info.WaitCount = int.Parse(Regex.Match(tmpWait, regRepC, RegexOptions.IgnoreCase).Groups["count"].Value);

                        string patrow = info.LastVector;
                        if (PatternGenBusiness.IsAddComment)
                        {
                            patrow += info.GetVectorInfo();
                            info.Update();
                        }
                        result.Add(patrow);
                    }
                    defV = tmpQ;

                }

                data = data.Replace("Wait", defV.ToString());
            }

            return data;
        }

        private static string ConvertAddressLSBFirst(string regAddress, int size)
        {
            string result = Regex.Match(regAddress, @"0x(?<data>\w+)", RegexOptions.IgnoreCase).Groups["data"].Value;

            result = string.Join(string.Empty, result.Select(c =>
                Convert.ToString(Convert.ToInt32(c.ToString(), 16), 2).PadLeft(4, '0')));
            if (result.Length < size)
            {
                result = result.PadLeft(size, '0');
            }

            char[] resultArr = result.ToCharArray();
            Array.Reverse(resultArr);
            return new string(resultArr);
        }

        public static void Write(List<string> patrows, string pattern, string subrName, int index, List<string> datas)
        {
            _ = new List<string>();
            var datalines = new List<string>();
            datalines.Add("digital_inst = hsdmq;");
            datalines.Add("opcode_mode = single;");
            datalines.Add("import tset tsetJTAG;");
            datalines.Add("instruments = {");
            datalines.Add("(JTAG_TDO):DigCap 1:lsb:serial:auto_trig_enable;");
            datalines.Add("(JTAG_TDI):DigSrc 1:lsb:serial;");
            datalines.Add("}");
            foreach (string data in datas)
            {
                if (data.StartsWith("//"))
                {
                    continue;
                }
                if (data.Contains("name"))
                {
                    string newName = pattern;
                    if (data.Contains("srm_vector"))
                    {
                        //newName = newName.Replace("");
                    }
                    else if (data.Contains("global"))
                    {
                        newName = subrName;
                    }

                    datalines.Add(data.Replace("name", newName));
                }
                else if (data.Contains("cnum"))
                {
                    string newindex = string.Format("c{0}", index);
                    datalines.Add(data.Replace("cnum", newindex));
                }
                else
                {
                    datalines.Add(data);
                }


            }
            foreach (string dataline in datalines)
            {
                patrows.Add(dataline);
            }
        }

        private static string AppendLogInfo(int recordindex, string addrORData, int maxNum)
        {
            string result;
            if (recordindex % maxNum == 0)
            {
                result = string.Format("{0} Start", addrORData);
            }
            else if (recordindex % maxNum == maxNum - 1)
            {
                result = string.Format("{0} End", addrORData);
            }
            else
            {
                return "";
            }

            return result;
        }

    }
}
