using System.Text.RegularExpressions;

using RF_PatternTool.PatternGen;

namespace RF_PatternTool.PatternStruct
{
    public class WritePattern
    {
        public List<string> RawDataList = new List<string>();
        private double _freq;
        public Dictionary<string, List<int>> SourceCaptureDictionary =
            new Dictionary<string, List<int>> { { "Source", new List<int>() }, { "Capture", new List<int>() } };
        public int PatBit = 32;

        static private string _jtagTdiSrcStr = "((JTAG_TDI):DigSrc = Send)";

        public WritePattern(double jtagFreq)
        {
            _freq = jtagFreq * 1e6;
        }
        //Read Raw data template
        public void ReadTemplate(string writeTemp)
        {
            int index = -1;
            List<string> lines = writeTemp.Replace("\r", "").Split('\n').ToList();
            foreach (string line in lines)
            {
                index++;
                if (line.Contains(">"))
                {
                    RawDataList.Add(line.Split('/')[0]);
                }
                else
                {
                    RawDataList.Add(line);
                }

                if (line.Contains(" D "))
                {
                    SourceCaptureDictionary["Source"].Add(index);
                }

                if (line.Contains(" V "))
                {
                    SourceCaptureDictionary["Capture"].Add(index);
                }
            }
        }

        // Send address
        // Setup achieve data position
        // Send address
        public void Write(List<string> patrows, BenchLogItem log, Register register, VectorInfo info, StreamWriter swlog)
        {
            string datainfo = log.Operation.Equals("Write", StringComparison.OrdinalIgnoreCase) || log.Operation.Equals("Reg_Write", StringComparison.OrdinalIgnoreCase) ? log.FieldVal : log.Default;
            string message = $"//OP:{log.Operation}. Address:{log.Address}. MSB:{log.MSB}. LSB:{log.LSB}. Data:{datainfo}.";
            try
            {
                List<BenchLogItem> srcinfos = log.SrcInfo;
                bool isRepeat = false;

                char[] dataArr = register.Data.ToCharArray();

                Array.Reverse(dataArr);
                string sourceDataStr = new string(dataArr); // Data
                string sourceAddressStr = ConvertAddressLSBFirst(register.Address, 32); // Address
                List<string> fieldList = register.GetFieldList(log);
                Writelog(swlog, register.Address, register.Data, register.SrcDic, log.RegData);
                int dataIndex = -1;
                //Assume 32 bits of data + 32 bits of Address
                //-> Assume (32 or 64) bits of data + 32 bits of Address
                string addressMsg = $"WriteAddr:{log.Address}";
                string writeDataMsg = string.IsNullOrEmpty(log.RegData) ? $"WriteData:{log.FieldVal}" : $"Reg_Data:{log.RegData}";

                for (int i = 0; i < RawDataList.Count; i++)
                {
                    string data = RawDataList[i];
                    if (data.Contains(_jtagTdiSrcStr))
                    {
                        dataIndex++;
                        if (dataIndex < PatBit) // send data information
                        {
                            string tDIData = sourceDataStr.Substring(0, 1);
                            sourceDataStr = sourceDataStr.Substring(1, sourceDataStr.Length - 1);
                            string dataMsg = AppendLogInfo(dataIndex, writeDataMsg, PatBit);
                            if (tDIData == "D")
                            {
                                patrows.Add(data);
                                data = RawDataList[++i];
                                info.CheckSourceInfo(fieldList[dataIndex]);
                                AddVectorRow(
                                    patrows,
                                    data + $"// Source Pin = JTAG_TDI sgmt{info.CurrentSourceSgmt} {fieldList[dataIndex]}",
                                    dataMsg,
                                    dataIndex,
                                    info);
                            }
                            else
                            {
                                data = RawDataList[++i];
                                data = data.Replace(" D ", " " + tDIData + " ");

                                AddVectorRow(patrows, data, dataMsg, dataIndex, info);

                                if (PatternGenBusiness.IsAddComment)
                                {
                                    patrows.Add(string.Format("// [{0}]", dataIndex));
                                }
                            }

                        }
                        else // send Address
                        {
                            i++; //Skip source syntax, use fixed value
                            string addrMsg = AppendLogInfo(dataIndex, addressMsg, 32);
                            data = RawDataList[i].Replace(" D ", " " + sourceAddressStr.Substring(0, 1) + " ");

                            AddVectorRow(patrows, data, addrMsg, dataIndex - PatBit, info);

                            if (PatternGenBusiness.IsAddComment)
                            {
                                patrows.Add(string.Format("// {1}address_csr[{0}]", dataIndex - PatBit, ""));
                            }
                            //address_csr[dataindex-32]
                            sourceAddressStr = sourceAddressStr.Substring(1, sourceAddressStr.Length - 1);
                        }

                    }
                    else
                    {
                        WriteOtherRow(patrows, data, i, log, info, message, ref isRepeat);
                    }
                }
                info.Clear();
            }
            catch (Exception)
            {
                ;
            }
        }

        private static void AddVectorRow(List<string> patrows, string data, string logMsg, int recordIndex, VectorInfo info)
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
            patrows.Add(patrow);
        }

        private void WriteOtherRow(List<string> patrows, string data, int rawIndex, BenchLogItem log, VectorInfo info, string message, ref bool isRepeat)
        {
            data = ExpandWaitRows(patrows, data, log, info);
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
                patrows.Add(patrow);

            }
            else
            {
                patrows.Add(data);
            }

        }

        private string ExpandWaitRows(List<string> patrows, string data, BenchLogItem log, VectorInfo info)
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
                        patrows.Add(tmpWait);
                        info.WaitCount =
                            int.Parse(
                                Regex.Match(tmpWait, regRepC, RegexOptions.IgnoreCase).Groups["count"].Value);

                        string patrow = info.LastVector;
                        if (PatternGenBusiness.IsAddComment)
                        {
                            patrow += info.GetVectorInfo();
                            info.Update();
                        }
                        patrows.Add(patrow);
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

        private static string AppendLogInfo(int recInd, string addrOrData, int patBit)
        {
            string result = "";

            if (recInd % patBit == 0)
            {
                result = $"{addrOrData} Start";
            }
            else if (recInd % patBit == patBit - 1)
            {
                result = $"{addrOrData} End";
            }

            return result;
        }

        private static void Writelog(StreamWriter sw, string address, string data, Dictionary<string, List<string>> dic, string reg_data)
        {
            if (data.Contains('D'))
            {
                //Addr0x44000274_M21_L20
                var regInfos_dic = new Dictionary<int, string>();
                string regMSBLSB = @"M(?<msb>\d+)_L(?<lsb>\d+)";
                foreach (KeyValuePair<string, List<string>> reg in dic)
                {
                    Match regSrcKey = Regex.Match(reg.Key, regMSBLSB, RegexOptions.IgnoreCase);
                    int msb = int.Parse(regSrcKey.Groups["msb"].Value);
                    int lsb = int.Parse(regSrcKey.Groups["lsb"].Value);

                    string regVal = reg.Value.Any() ? reg.Value.Last() : "";
                    regInfos_dic.Add(lsb, $"{regVal}#{lsb}_{msb - lsb + 1}");
                }
                IEnumerable<string> regInfos = regInfos_dic.OrderBy(p => p.Key).Select(p => p.Value);
                sw.WriteLine($"RegisterWrite_DigSrc,{address},{reg_data},{string.Join(",", regInfos)}");
            }
            else
            {
                sw.WriteLine($"RegisterWrite,{address},{reg_data}");
            }
        }
    }
}
