using System.Diagnostics;
using System.Text.RegularExpressions;

namespace RF_PatternTool.PatternGen
{
    public class BenchLogFile
    {
        public string LogFile;
        public string DateCode;
        public string Version;
        public string Interface;
        public string PatternSubName;
        public string Type;
        public bool IsReadCapTrim;              //FullSweep
        //public bool IsNeedOverWrite = false;     //HistoryChecking
        public Dictionary<string, string> WriteSrcRows = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<string> Inits = new List<string>();
        public int Offset;
        public List<BenchLogItem> Logs;
        public List<RFLogError> Errors;
        public List<string> RawDatas = new List<string>();
        public static string HeaderOperation = "Operation";
        public static string HeaderInterface = "Interface";
        public static string HeaderAddress = "Address";
        public static string HeaderMSB = "MSB";
        public static string HeaderLSB = "LSB";
        public static string HeaderRegFieldName = "RegFieldName";
        public static string HeaderData = "Data";
        public static string HeaderFieldVal = "FieldVal";
        public static string HeaderValue = "Value";
        public static string HeaderDefault = "Default";
        public static string HeaderFuseName = "eFuseName";
        public static string HeaderPostProcess = "PostProcess";
        public static string HeaderTestName = "TestName";
        public static string HeaderLowLimit = "LimitLo";
        public static string HeaderHighLimit = "LimitHi";
        public static string HeaderUnit = "Units";
        public static string HeaderRegData = "RegData";

        public int IndexOperation = -1;
        public int IndexInterface = -1;
        public int IndexAddress = -1;
        public int IndexMSB = -1;
        public int IndexLSB = -1;
        public int IndexRegFieldName = -1;
        public int IndexData = -1;
        public int IndexDefault = -1;
        public int IndexFuseName = -1;
        public int IndexPostProcess = -1;
        public int IndexTestName = -1;
        public int IndexLowLimit = -1;
        public int IndexHighLimit = -1;
        public int IndexUnit = -1;
        public int IndexRegData = -1;

        public BenchLogFile()
        {
            Logs = new List<BenchLogItem>();
        }


        public void GetHeader(string log)
        {
            string[] datas = log.Split(',');
            for (int i = 0; i < datas.Count(); i++)
            {
                string value = datas[i].Trim().Replace("_", "").Replace(" ", "");
                if (value.Equals(HeaderOperation, StringComparison.OrdinalIgnoreCase))
                { IndexOperation = i; }
                else if (value.Equals(HeaderInterface, StringComparison.OrdinalIgnoreCase))
                { IndexInterface = i; }
                else if (value.Equals(HeaderAddress, StringComparison.OrdinalIgnoreCase))
                { IndexAddress = i; }
                else if (value.Equals(HeaderMSB, StringComparison.OrdinalIgnoreCase))
                { IndexMSB = i; }
                else if (value.Equals(HeaderLSB, StringComparison.OrdinalIgnoreCase))
                { IndexLSB = i; }
                else if (value.Equals(HeaderRegFieldName, StringComparison.OrdinalIgnoreCase))
                { IndexRegFieldName = i; }
                else if (value.Equals(HeaderData, StringComparison.OrdinalIgnoreCase) ||
                    value.Equals(HeaderFieldVal, StringComparison.OrdinalIgnoreCase) ||
                    value.Equals(HeaderValue, StringComparison.OrdinalIgnoreCase))
                { IndexData = i; }
                else if (value.Equals(HeaderDefault, StringComparison.OrdinalIgnoreCase))
                { IndexDefault = i; }
                else if (value.Equals(HeaderFuseName, StringComparison.OrdinalIgnoreCase))
                { IndexFuseName = i; }
                else if (value.Equals(HeaderPostProcess, StringComparison.OrdinalIgnoreCase))
                { IndexPostProcess = i; }
                else if (value.Equals(HeaderTestName, StringComparison.OrdinalIgnoreCase))
                { IndexTestName = i; }
                else if (value.Equals(HeaderLowLimit, StringComparison.OrdinalIgnoreCase))
                { IndexLowLimit = i; }
                else if (value.Equals(HeaderHighLimit, StringComparison.OrdinalIgnoreCase))
                { IndexHighLimit = i; }
                else if (value.Equals(HeaderUnit, StringComparison.OrdinalIgnoreCase))
                { IndexUnit = i; }
                else if (value.Equals(HeaderRegData, StringComparison.OrdinalIgnoreCase))
                { IndexRegData = i; }
            }
        }

        private void SetHeaderIndex(string value, int i)
        {
            if (value.Equals(HeaderOperation, StringComparison.OrdinalIgnoreCase))
            { IndexOperation = i; }
            else if (value.Equals(HeaderInterface, StringComparison.OrdinalIgnoreCase))
            { IndexInterface = i; }
            else if (value.Equals(HeaderAddress, StringComparison.OrdinalIgnoreCase))
            { IndexAddress = i; }
            else if (value.Equals(HeaderMSB, StringComparison.OrdinalIgnoreCase))
            { IndexMSB = i; }
            else if (value.Equals(HeaderLSB, StringComparison.OrdinalIgnoreCase))
            { IndexLSB = i; }
            else if (value.Equals(HeaderRegFieldName, StringComparison.OrdinalIgnoreCase))
            { IndexRegFieldName = i; }
            else if (value.Equals(HeaderData, StringComparison.OrdinalIgnoreCase) ||
                value.Equals(HeaderFieldVal, StringComparison.OrdinalIgnoreCase) ||
                value.Equals(HeaderValue, StringComparison.OrdinalIgnoreCase))
            { IndexData = i; }
            else if (value.Equals(HeaderDefault, StringComparison.OrdinalIgnoreCase))
            { IndexDefault = i; }
            else if (value.Equals(HeaderFuseName, StringComparison.OrdinalIgnoreCase))
            { IndexFuseName = i; }
            else if (value.Equals(HeaderPostProcess, StringComparison.OrdinalIgnoreCase))
            { IndexPostProcess = i; }
            else if (value.Equals(HeaderTestName, StringComparison.OrdinalIgnoreCase))
            { IndexTestName = i; }
            else if (value.Equals(HeaderLowLimit, StringComparison.OrdinalIgnoreCase))
            { IndexLowLimit = i; }
            else if (value.Equals(HeaderHighLimit, StringComparison.OrdinalIgnoreCase))
            { IndexHighLimit = i; }
            else if (value.Equals(HeaderUnit, StringComparison.OrdinalIgnoreCase))
            { IndexUnit = i; }
            else if (value.Equals(HeaderRegData, StringComparison.OrdinalIgnoreCase))
            { IndexRegData = i; }
        }

        private List<string> GetMissingHeaders()
        {
            var missingHeaders = new List<string>();
            if (IndexOperation == -1)
            {
                missingHeaders.Add(HeaderOperation);
            }

            if (IndexInterface == -1)
            {
                missingHeaders.Add(HeaderInterface);
            }

            if (IndexAddress == -1)
            {
                missingHeaders.Add(HeaderAddress);
            }

            if (IndexMSB == -1)
            {
                missingHeaders.Add(HeaderMSB);
            }

            if (IndexLSB == -1)
            {
                missingHeaders.Add(HeaderLSB);
            }

            if (IndexRegFieldName == -1)
            {
                missingHeaders.Add(HeaderRegFieldName);
            }

            if (IndexData == -1)
            {
                missingHeaders.Add(HeaderData);
            }
            //if (IndexDefault == -1)
            //    missingHeaders.Add(HeaderDefault);
            if (IndexFuseName == -1)
            {
                missingHeaders.Add(HeaderFuseName);
            }

            if (IndexPostProcess == -1)
            {
                missingHeaders.Add(HeaderPostProcess);
            }

            if (IndexTestName == -1)
            {
                missingHeaders.Add(HeaderTestName);
            }

            if (IndexLowLimit == -1)
            {
                missingHeaders.Add(HeaderLowLimit);
            }

            if (IndexHighLimit == -1)
            {
                missingHeaders.Add(HeaderHighLimit);
            }

            if (IndexUnit == -1)
            {
                missingHeaders.Add(HeaderUnit);
            }

            if (IndexRegData == -1)
            {
                missingHeaders.Add(HeaderRegData);
            }

            return missingHeaders;
        }

        public List<string> GetHeaderAndCheck(string log)
        {
            string[] datas = log.Split(',');
            for (int i = 0; i < datas.Count(); i++)
            {
                string value = datas[i].Trim().Replace("_", "").Replace(" ", "");
                SetHeaderIndex(value, i);
            }
            //check header
            #region check header
            return GetMissingHeaders();
            //
            #endregion
        }

        public BenchLogFile Copy()
        {
            var log = new BenchLogFile();
            log.Offset = Offset;
            log.IndexOperation = IndexOperation;
            log.IndexAddress = IndexAddress;
            log.IndexInterface = IndexInterface;
            log.IndexMSB = IndexMSB;
            log.IndexLSB = IndexLSB;
            log.IndexRegFieldName = IndexRegFieldName;
            log.IndexData = IndexData;
            log.IndexRegData = IndexRegData;
            log.IndexDefault = IndexDefault;
            log.IndexFuseName = IndexFuseName;

            log.IndexPostProcess = IndexPostProcess;
            log.IndexTestName = IndexTestName;
            log.IndexLowLimit = IndexLowLimit;
            log.IndexHighLimit = IndexHighLimit;
            log.IndexUnit = IndexUnit;

            foreach (string data in RawDatas)
            {
                log.RawDatas.Add(data);
            }
            RawDatas.Clear();

            return log;
        }
    }

    public class CaptureInfo
    {
        public bool IsFullCap = false;
        public string TestName = "";
        public string Address = "";
        public int LSB = -1;
        public int MSB = -1;
        public string PostProcess = "";
        public string LimitHight = "";
        public string LimitLow = "";
        public string Unit = "";
        public string FuseName = "";
    }

    [DebuggerDisplay("{RawData}")]
    public class BenchLogItem
    {
        public int RowNum = 0;
        public int Offset = 0;
        public string Operation = "";
        public string Interface = "";
        public string Address = "";
        public int MSB = -1;
        public int LSB = -1;
        public string RegFieldName = "";
        public string FieldVal = "";
        public string Default = "";
        public string FuseName = "";
        public double Wait = 0.0;
        public string RawData = "";
        public List<string> RawDatas = new List<string>();
        public string PostProcess = "";
        public string TestName = "";
        public string LoLimit = "";
        public string HighLimit = "";
        public string Units = "";
        //public string patName = "";
        //public string patVer = "";
        public string RegData = "";
        public bool IsFullSweep = false;
        public List<CaptureInfo> Capinfos = new List<CaptureInfo>();
        public List<BenchLogItem> SrcInfo = new List<BenchLogItem>();

        public BenchLogItem(string log, BenchLogFile logInfo, int row, bool isfullSweep, ref bool isFirst)
        {
            IsFullSweep = isfullSweep;
            List<string> datas = log.Trim('"').Split(',').ToList();
            RawData = log;
            RowNum = row;
            if (log.Contains("HSC_WL5GADC1CAL_JTAG-TDO_MEASC_X_X_ERROR7-0-I_X_NV_VXXX"))
            {
                ;
            }
            if (log.Contains("") && !isFirst)
            {
                int quoteNum = 0;
                var data_new = new List<string>();
                bool isMerge = false;
                int rem = 0;
                foreach (string data in datas)
                {
                    quoteNum = quoteNum + data.Count(p => p == '"');
                    Math.DivRem(quoteNum, 2, out rem);
                    if (isMerge)
                    {
                        string line = data_new.Count == 0 ? data : data_new.Last() + "," + data;
                        data_new.Remove(data_new.Last());
                        data_new.Add(line.Trim('"'));
                        if (rem == 0)
                        {
                            isMerge = false;
                            quoteNum = 0;
                        }
                    }
                    else
                    {
                        data_new.Add(data.Trim('"'));
                        if (rem != 0)
                        {
                            isMerge = true;
                        }
                    }
                    datas = data_new;
                }

            }
            else if (isFirst)
            {
                isFirst = false;
            }

            RawDatas.AddRange(datas);
            for (int i = 0; i < datas.Count(); i++)
            {
                SetPrimaryFieldValue(datas, logInfo, i);
                SetSecondaryFieldValue(datas, logInfo, i);
            }
            if (logInfo.IndexRegData != -1)
            {
                //patName = RegData;
                //patVer = RegFieldName;
                RegFieldName = MSB.ToString();
            }

            //if (string.IsNullOrEmpty(Default))
            //    Default = Data;
            if (Operation.Equals("READ_CAP", StringComparison.OrdinalIgnoreCase) ||
                Operation.Equals("SPAC_MEM_READ", StringComparison.OrdinalIgnoreCase))
            {
                var capinfo = new CaptureInfo();
                capinfo.Address = Address;
                capinfo.TestName = TestName;
                capinfo.LSB = LSB;
                capinfo.MSB = MSB;
                capinfo.PostProcess = PostProcess;
                capinfo.LimitHight = HighLimit;
                capinfo.LimitLow = LoLimit;
                capinfo.Unit = Units;
                capinfo.FuseName = FuseName;
                capinfo.IsFullCap = IsFullSweep;
                Capinfos.Add(capinfo);
            }

            if (Regex.IsMatch(Operation, "Write", RegexOptions.IgnoreCase))
            {
                SrcInfo.Add(this);
            }
        }

        private void SetPrimaryFieldValue(List<string> datas, BenchLogFile logInfo, int i)
        {
            if (i == logInfo.IndexOperation)
            {
                Operation = datas[i].Trim();
            }

            if (i == logInfo.IndexInterface)
            {
                Interface = datas[i].Trim();
            }

            if (i == logInfo.IndexAddress)
            {
                Address = datas[i].Trim();
                if (!Operation.Equals("PATNAME", StringComparison.OrdinalIgnoreCase) &&
                    !Operation.Equals("OTP_WRITE", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrEmpty(Address) && !Regex.IsMatch(Address, "^0x", RegexOptions.IgnoreCase))
                {
                    Address = "0x" + Address;
                }
            }
            if (i == logInfo.IndexMSB)
            {
                int.TryParse(datas[i].Trim(), out MSB);
            }
            if (i == logInfo.IndexLSB)
            {
                int.TryParse(datas[i].Trim(), out LSB);
            }
            if (i == logInfo.IndexRegFieldName)
            {
                RegFieldName = datas[i].Trim();
            }
            if (i == logInfo.IndexData)
            {
                FieldVal = datas[i].Trim();

                //if (!string.IsNullOrEmpty(Data)&& !Regex.IsMatch(Data, "^0x", RegexOptions.IgnoreCase))
                //Data = "0x" + Data;
            }
        }

        private void SetSecondaryFieldValue(List<string> datas, BenchLogFile logInfo, int i)
        {
            if (i == logInfo.IndexDefault)
            {
                Default = datas[i].Trim();
                if (!string.IsNullOrEmpty(Default) && !Regex.IsMatch(Default, "^0x", RegexOptions.IgnoreCase))
                {
                    Default = "0x" + Default;
                }
            }
            if (i == logInfo.IndexFuseName)
            {
                FuseName = datas[i].Trim();
            }
            if (i == logInfo.IndexRegData)
            {
                RegData = datas[i].Trim();
            }

            if (i == logInfo.IndexPostProcess)
            {
                PostProcess = datas[i].Trim();
            }

            if (i == logInfo.IndexTestName)
            {
                TestName = datas[i].Trim();
            }

            if (i == logInfo.IndexLowLimit)
            {
                LoLimit = datas[i].Trim();
            }

            if (i == logInfo.IndexHighLimit)
            {
                HighLimit = datas[i].Trim();
            }

            if (i == logInfo.IndexUnit)
            {
                Units = datas[i].Trim();
            }
        }

        public string GetAddressInfo()
        {
            return string.Format("{0}#{1}#{2}", Address, LSB, MSB).ToUpper();
        }


        public string GetSpecialData()
        {
            string result = Interface;
            if (Interface.Contains("\""))
            {
                string regMsg = "\"" + @"(?<msg>.*)" + "\"";
                string tmp = Regex.Match(RawData.Replace(Operation, ""), regMsg, RegexOptions.IgnoreCase).Groups["msg"].Value;
                if (!string.IsNullOrEmpty(tmp))
                {
                    result = tmp;
                }
            }
            return result;
        }
    }
}
