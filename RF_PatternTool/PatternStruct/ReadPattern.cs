using System.Text.RegularExpressions;

using RF_PatternTool.PatternGen;

using RfLib.Dvdc.Reader.CapturePostProcess;

using RFPatternTool;

namespace RF_PatternTool.PatternStruct
{
    public class ReadPattern
    {
        public List<string> RawDataList = new List<string>();
        public int StoreIndex = -1;
        public string PatSubName = "";
        private double _freq;
        public Dictionary<string, List<int>> SourceCaptureDictionary =
            new Dictionary<string, List<int>> { { "Source", new List<int>() }, { "Capture", new List<int>() } };
        public int PatBit = 32;

        static private string _jtagTdiSrcStr = "((JTAG_TDI):DigSrc = Send)";
        static private string _jtagTdoCapStr = "((JTAG_TDO):DigCap = Store)";

        public ReadPattern(double jtagFreq)
        {
            _freq = jtagFreq * 1e6;
        }
        public void ReadTemplate(string readTemp)
        {
            int index = -1;
            List<string> lines = readTemp.Replace("\r", "").Split('\n').ToList();
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
        public void Write(List<string> patrows, BenchLogItem log, Register register, VectorInfo info, ref List<PostProcessSheetRow> cppList, StreamWriter swlog, bool isTrimStart, PatternGenItem itemInfo)
        {
            bool isSPAC = Regex.IsMatch(log.Operation, "spac_mem_read", RegexOptions.IgnoreCase);
            int bwtimes = isSPAC ? Convert.ToInt32(log.RegData, 10) : 1;

            string message = $"//OP:{log.Operation}. Address:{log.Address}. Field:{log.RegFieldName.Split('.').Last()}.";
            bool isRepeat = false;
            string sourceAddressStr = ConvertAddressLSBFirst(register.Address, 32) + ConvertAddressLSBFirst(register.Address, 32);// Address
            List<string> fieldList = register.GetFieldList(log);
            char[] dataList = ("".PadLeft(register.Data.Length, 'X')).ToCharArray();
            int patAddIndex = -1;
            int patDataIndex = -1;
            //Assume 32 bits of data + 32 bits of Address
            //-> Assume 32 bits of Address + (32 or 64) bits of data

            int shiftbit = 0;
            int addrQuot = Convert.ToInt32(log.Address, 16) / 4;
            bool isNeedShift = PatBit == 64 && addrQuot % 2 != 0;
            if (isNeedShift)
            {
                shiftbit = 32;
            }

            string addressMsg = $"ReadAddr:{log.Address}";

            string dataReadMsg = $"ReadData:{log.FieldVal}";
            if (log.Capinfos.Count > 0)
            {
                foreach (CaptureInfo capinfo in log.Capinfos)
                {
                    for (int i = capinfo.LSB; i <= capinfo.MSB; i++)
                    {
                        dataList[i + shiftbit] = 'V';
                    }
                    var cpp = new PostProcessSheetRow();
                    AnalyzeCalc(cpp, capinfo, isTrimStart, bwtimes);

                    cpp.BitWidth = (capinfo.MSB - capinfo.LSB + 1).ToString();
                    cpp.TestName = $"Addr_{log.Address}_L{capinfo.LSB}_M{capinfo.MSB}";
                    cppList.Add(cpp);
                    Writelog(swlog, register.Address, dataList, capinfo, isNeedShift);
                }
            }
            else
            {
                string binValue = Convert.ToString((long)(log.FieldVal.StartsWith("0x") ?
                        Convert.ToUInt64(log.FieldVal.Replace("0x", ""), 16) :
                        Convert.ToUInt64(log.FieldVal)), 2);

                binValue = binValue.PadLeft(PatBit, '0').Replace("0", "L").Replace("1", "H");
                // change to compare mode

                //reverse
                char[] revArr = binValue.ToCharArray();
                Array.Reverse(revArr);
                int index = 0;
                for (int i = log.LSB; i <= log.MSB; i++)
                {
                    dataList[i + shiftbit] = revArr[index];
                    index++;
                }
                Writelog(swlog, register.Address, dataList, null, isNeedShift);
            }

            for (int i = 0; i < RawDataList.Count; i++)
            {
                string data = RawDataList[i];

                if (data.Contains(_jtagTdiSrcStr))
                {
                    //address
                    patAddIndex++;
                    _ = Math.DivRem(patAddIndex, 32, out patAddIndex);
                    i++; //Skip source syntax, use fixed value
                    data = RawDataList[i].Replace(" D ", " " + sourceAddressStr.Substring(0, 1) + " ");

                    string addrMsg = AppendLogInfo(patAddIndex, addressMsg, 32);
                    AddVectorRow(patrows, data, addrMsg, patAddIndex, info);
                    if (PatternGenBusiness.IsAddComment)
                    {
                        patrows.Add(string.Format("// address_csr[{0}]", patAddIndex)); //address_csr[dataindex-32]
                    }

                    sourceAddressStr = sourceAddressStr.Substring(1, sourceAddressStr.Length - 1); //remove LSB data
                }
                else if (data.Contains(_jtagTdoCapStr))
                {
                    patDataIndex++;
                    _ = Math.DivRem(patDataIndex, PatBit, out patDataIndex);

                    string dataMsg = AppendLogInfo(patDataIndex, dataReadMsg, PatBit);
                    if (dataList[patDataIndex] == 'V')
                    {
                        patrows.Add(data);
                        data = RawDataList[++i];
                        info.CheckSourceInfo(fieldList[patDataIndex]);

                        AddVectorRow(
                            patrows,
                            data + $"// Capture Pin = JTAG_TDO sgmt{info.CurrentSourceSgmt} {fieldList[patDataIndex]}",
                            dataMsg,
                            patDataIndex,
                            info);
                    }
                    else
                    {
                        try
                        {
                            data = RawDataList[++i];
                            data = data.Replace(" V ", " " + dataList[patDataIndex] + " ");

                            AddVectorRow(patrows, data, dataMsg, patDataIndex, info);
                            if (PatternGenBusiness.IsAddComment)
                            {
                                patrows.Add(string.Format("// [{0}]", patDataIndex));
                            }
                        }
                        catch (Exception)
                        {
                            ;
                        }
                    }
                }
                else
                {
                    WriteOtherRow(patrows, data, i, log, info, message, ref isRepeat);
                }
            }
            info.Clear();

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
                if (data.StartsWith("//"))
                {
                    string regData = @"data\[(?<num>\d+)\]";
                    if (data.ToLower().Contains("data") && Regex.IsMatch(data, regData, RegexOptions.IgnoreCase))
                    {
                        _ = int.Parse(Regex.Match(data, regData, RegexOptions.IgnoreCase).Groups["num"].Value);
                        patrows.Add("// ");
                    }
                    else
                    {
                        ;
                    }

                }
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
                        info.WaitCount = int.Parse(Regex.Match(tmpWait, regRepC, RegexOptions.IgnoreCase).Groups["count"].Value);

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

        private void AnalyzeCalc(PostProcessSheetRow cpp, CaptureInfo capInfo, bool isTrimStart, int bwtimes)
        {
            int i_name = -1;
            var postCalcItem = new PostCalcInfo();
            postCalcItem.CalcTestName = capInfo.TestName;
            if (string.IsNullOrEmpty(capInfo.LimitHight))
            {
                postCalcItem.HiLimit = "NA";
            }
            else
            {
                postCalcItem.HiLimit = capInfo.LimitHight;
            }

            postCalcItem.HiLimit = postCalcItem.HiLimit + capInfo.Unit;

            if (string.IsNullOrEmpty(capInfo.LimitLow))
            {
                postCalcItem.LowLimit = "NA";
            }
            else
            {
                postCalcItem.LowLimit = capInfo.LimitLow;
            }

            postCalcItem.LowLimit = postCalcItem.LowLimit + capInfo.Unit;

            postCalcItem.Bit = (capInfo.MSB - capInfo.LSB + 1) * bwtimes;
            var calcs = new List<string>();

            string capInPosPro = string.IsNullOrEmpty(capInfo.PostProcess) ?
                $"StoreName:{PatSubName}_Store{StoreIndex}" : capInfo.PostProcess;

            bool withCalc = false;

            foreach (string subline in capInPosPro.Trim(',').Replace("=", ":").Split(';'))
            {
                i_name++;
                if (subline.Contains(':'))
                {
                    string key = subline.Split(':')[0].Trim().ToLower();
                    string value = subline.Split(':')[1].Trim();
                    switch (key)
                    {
                        case "calc":
                            postCalcItem.CalcEquation = value;
                            withCalc = true;
                            break;
                        case "storename":
                            cpp.StoreName = value;
                            calcs.Add(string.Format("rffw_calc_RAW({0})", value));
                            break;
                        case "calcstorename":
                            postCalcItem.CalcStoreName = value;
                            if (isTrimStart && !capInfo.IsFullCap)
                            {
                                postCalcItem.CalcStoreName += "_CapTrimData";
                            }

                            break;
                    }
                }
                else
                {
                    calcs.Add(subline);
                }
            }

            capInfo.PostProcess = string.IsNullOrEmpty(capInfo.PostProcess) ?
                string.Format("StoreName:{0}_Store{1};Calc:rffw_calc_RAW({0}_Store{1})", PatSubName, StoreIndex++) :
                withCalc ?
                capInfo.PostProcess : capInfo.PostProcess + $";Calc:rffw_calc_RAW({cpp.StoreName})";

            postCalcItem.CalcEquation = string.IsNullOrEmpty(postCalcItem.CalcEquation) ?
                string.Join(";", calcs) : postCalcItem.CalcEquation;

            string[] csNames = postCalcItem.CalcStoreName.Split(',');
            string[] tNames = capInfo.TestName.Split(',');

            AddPostCalcs(cpp, capInfo, postCalcItem, csNames, tNames);
        }

        private static void AddPostCalcs(PostProcessSheetRow cpp, CaptureInfo capInfo, PostCalcInfo postCalcItem, string[] csNames, string[] tNames)
        {
            if (tNames.Length > 1 && tNames.Length != csNames.Length)
            {
                ;
            }
            else if (tNames.Length == 1 && csNames.Length == 1)
            {
                cpp.PostCalcs.Add(new PostCalcInfo(
                    postCalcItem.CalcEquation,
                    capInfo.TestName,
                    postCalcItem.CalcStoreName,
                    postCalcItem.LowLimit,
                    postCalcItem.HiLimit,
                    postCalcItem.Bit));
            }
            else
            {
                for (int i = 0; i < csNames.Length; i++)
                {
                    string calceqa = i == 0 ? postCalcItem.CalcEquation : "";

                    string[] splitCNName = tNames.Length == 1 ?
                        tNames[0].Split('_') : tNames[i].Split('_');
                    if (tNames.Length == 1)
                    {
                        if (splitCNName.Length >= 3 && !string.IsNullOrEmpty(csNames[i]))
                        {
                            splitCNName[3] = csNames[i].Replace("_", "").ToUpper();
                        }
                    }

                    cpp.PostCalcs.Add(new PostCalcInfo(
                        calceqa,
                        string.Join("_", splitCNName),
                        csNames[i],
                        postCalcItem.LowLimit,
                        postCalcItem.HiLimit,
                        i == 0 ? postCalcItem.Bit : 0));
                }
            }
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

        private void Writelog(StreamWriter sw, string address, char[] data, CaptureInfo capinfo, bool isNeedShift)
        {
            /*
             * RegisterWrite_DigSrc	0x44000274	 0x100803c0	 lpclk_best_cap#2_5
             * RegisterWrite	 0x4400919c	0xf
             */

            Array.Reverse(data);

            _ = string.Join("", data);
            if (capinfo == null)
            {
                string mask = (new string(data)).Replace("H", "1").Replace("L", "1").Replace("X", "0");
                string regVal = @"(?<val>[HL]+)";
                string value = Regex.Match(new string(data), regVal, RegexOptions.IgnoreCase).Groups["val"].Value.Replace("H", "1").Replace("L", "0");

                string maskX = $"0x{(data.Length == 32 ? Convert.ToInt32(mask, 2).ToString("X") : Convert.ToInt64(mask, 2).ToString("X"))}";
                if (maskX.EndsWith("00000000") && isNeedShift)
                {
                    maskX = maskX.Substring(0, maskX.Length - 8);
                }

                sw.WriteLine($"RegisterCompare," +
                    $"{address}," +
                    $"0x{(data.Length == 32 ? Convert.ToInt32(value, 2).ToString("X") : Convert.ToInt64(value, 2).ToString("X"))}," +
                    $"{maskX}");
            }
            else
            {
                string mask = "";
                for (int i = 0; i <= PatBit - 1; i++)
                {
                    if (i <= capinfo.MSB && i >= capinfo.LSB)
                    {
                        mask = "1" + mask;
                    }
                    else
                    {
                        mask = "0" + mask;
                    }
                }

                sw.WriteLine($"RegisterRead,{address}," +
                    $"0x{(data.Length == 32 ? Convert.ToInt32(mask, 2).ToString("X") : Convert.ToInt64(mask, 2).ToString("X"))}," +
                    $"{capinfo.TestName}");
            }
            Array.Reverse(data);
        }
    }
}
