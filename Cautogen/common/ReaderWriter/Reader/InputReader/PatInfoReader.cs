using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.common.ReaderWriter.Reader.InputDataBase;

namespace Cautogen.common.ReaderWriter.Reader.InputReader
{
    public class PatInfoReader : TextInputReader
    {
        private List<HardIpReference> _hardIpInfo = new List<HardIpReference>();

        /* properties */
        public Dictionary<string, string> PinUnderlineDict;

        /* constructor */
        public PatInfoReader(string filePath) : base(filePath)
        {
            PinUnderlineDict = new Dictionary<string, string>();
        }
        public new List<HardIpReference> Read()
        {
            if (IsFileExist())
            {
                FileStream fileStream = null;

                try
                {
                    fileStream = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
                    using (var textReader = new StreamReader(fileStream))
                    {
                        fileStream = null;
                        _Read(textReader);
                    }
                }
                finally
                {
                    if (fileStream != null)
                    {
                        fileStream.Dispose();
                    }
                }
            }
            return _hardIpInfo;
        }

        /* methods */
        protected override void _Read(StreamReader textReader)
        {
            if (!File.Exists(FilePath))
            {
                return;
            }

            var mHardIpInfo = new List<HardIpReference>();
            var nHardIpInfo = new HardIpReference();

            int count = 0;
            bool start = false;
            int rowNum = 0;
            try
            {
                string param;
                while ((param = textReader.ReadLine()) != null)
                {
                    rowNum++;
                    if (param.StartsWith("GenericPat:=", StringComparison.CurrentCultureIgnoreCase))
                    {
                        HandleSeparator(mHardIpInfo, ref nHardIpInfo, ref count, ref start, rowNum);
                    }
                    if (!start)
                    {
                        continue;
                    }

                    param = param.Replace(" ", "");
                    param = param.Replace("::", "&");
                    string[] arrstr = param.Split(':');
                    int index = param.IndexOf(":", StringComparison.Ordinal);
                    string tmpstr = index != -1 ? param.Substring(index + 1, param.Length - index - 1)
                                         : param;
                    /*if (Regex.IsMatch(tmpstr, @"[^a-zA-Z0-9_+,=\<\>):&\|.]+"))
                    {
                        MessageWriter.WriteMessage("Error : " + param + "Contain Illegal Character",
                            MessageLevel.Error);
                        nHardIpInfo.IllegalChar = true;
                    }*/
                    ParseLine(arrstr, nHardIpInfo);
                }

                FlushPending(mHardIpInfo, nHardIpInfo);
            }
            catch (Exception e)
            {
                throw new Exception("Extracting patInfo file failed at row: " + rowNum + " " + nHardIpInfo.Payload + "  ---\r\n " + "\r\n" + e.Message);
            }
            _hardIpInfo = mHardIpInfo;
        }

        private void HandleSeparator(List<HardIpReference> mHardIpInfo, ref HardIpReference nHardIpInfo, ref int count, ref bool start, int rowNum)
        {
            if (count == 0)
            {
                start = true;
            }
            else
            {
                HardIpReference current = nHardIpInfo;
                //prevent duplicate add
                if (mHardIpInfo.FindIndex(x => string.Equals(x.Payload + x.Version, current.Payload + current.Version, StringComparison.CurrentCultureIgnoreCase)) == -1)
                {
                    if (!current.IllegalChar)
                    {
                        try
                        {
                            current.SeqInfo = ExtractMeasSeqInfo(current);
                            mHardIpInfo.Add(current);
                        }
                        catch (Exception e)
                        {
                            string outString = string.Format("Extracting patInfo file failed at row: " + (rowNum - 1) + " " + current.Payload + "  ---\r\n " + "\r\n" + e.Message);
                        }
                    }
                }
                nHardIpInfo = new HardIpReference();
            }
            count++;
        }

        private void FlushPending(List<HardIpReference> mHardIpInfo, HardIpReference nHardIpInfo)
        {
            if (mHardIpInfo.FindIndex(x => string.Equals(x.Payload + x.Version, nHardIpInfo.Payload + nHardIpInfo.Version, StringComparison.CurrentCultureIgnoreCase)) == -1)
            {
                if (!nHardIpInfo.IllegalChar)
                {
                    nHardIpInfo.SeqInfo = ExtractMeasSeqInfo(nHardIpInfo);
                    mHardIpInfo.Add(nHardIpInfo);
                }
            }
        }

        private void ParseLine(string[] arrstr, HardIpReference nHardIpInfo)
        {
            if (ParseMetaField(arrstr, nHardIpInfo))
            {
                return;
            }
            ParseDataField(arrstr, nHardIpInfo);
        }

        private static bool ParseMetaField(string[] arrstr, HardIpReference nHardIpInfo)
        {
            string key = arrstr[0].ToUpper();
            if (key.Contains("GENERICPAT"))
            {
                nHardIpInfo.Payload = arrstr[1].ToUpper().Replace("=", "");
                return true;
            }
            if (key.Contains("VERSION"))
            {
                nHardIpInfo.Version = arrstr[1].ToUpper().Replace("=", "");
                return true;
            }
            if (key.Contains("TSET"))
            {
                nHardIpInfo.TimeSet = arrstr[1];
                return true;
            }
            if (key == "CAPPINNAME")
            {
                nHardIpInfo.CapPinName = arrstr[1];
                return true;
            }
            if (key == "CAPBIT")
            {
                nHardIpInfo.CapBit = Convert.ToInt32(arrstr[1]);
                return true;
            }
            if (key.Contains("CAPBITSTR"))
            {
                nHardIpInfo.CapBitStr = arrstr[1];
                return true;
            }
            if (key.Contains("CAPBITNAME"))
            {
                nHardIpInfo.CapBitName = arrstr[1];
                return true;
            }
            if (key == "SENDPINNAME")
            {
                nHardIpInfo.SendPinName = arrstr[1];
                return true;
            }
            if (key.Contains("SENDBITNAME"))
            {
                nHardIpInfo.SendBitName = arrstr[1].Trim();
                return true;
            }
            if (key == "SENDBIT")
            {
                nHardIpInfo.SendBit = Convert.ToInt32(arrstr[1]);
                return true;
            }
            return false;
        }

        private void ParseDataField(string[] arrstr, HardIpReference nHardIpInfo)
        {
            string key = arrstr[0].ToUpper();
            if (key.Contains("SENDBITSTR"))
            {
                nHardIpInfo.SendBitStr = arrstr[1];
                if (RemoveDummy(nHardIpInfo.SendBitStr) == "")
                {
                    return;
                }
                string tmpstr = ExtractDigSrcInfo(RemoveDummy(nHardIpInfo.SendBitStr));
                string[] arrstr1 = tmpstr.Split(';');
                nHardIpInfo.DigSrcDataWidth = arrstr1[0];
                nHardIpInfo.DigSrcAssign = arrstr1[1];
                return;
            }
            if (key.Contains("MEASSEQ"))
            {
                nHardIpInfo.MeasSeqStr = arrstr[1].ToUpper();
                return;
            }
            ParseMeasField(arrstr, nHardIpInfo);
        }

        private void ParseMeasField(string[] arrstr, HardIpReference nHardIpInfo)
        {
            switch (arrstr[0].ToUpper())
            {
                case "F":
                    arrstr[1] = PinConvertMapping(arrstr[1].ToUpper());
                    arrstr[1] = arrstr[1].Replace("&", "::");
                    nHardIpInfo.MeasFStr = arrstr[1];
                    AddTokens(arrstr[1], nHardIpInfo.MeasFPinList);
                    break;
                case "I":
                    arrstr[1] = PinConvertMapping(arrstr[1].ToUpper());
                    nHardIpInfo.MeasIStr = arrstr[1];
                    AddTokensWithVddContainsSplit(arrstr[1], nHardIpInfo.MeasIPowerPinList, nHardIpInfo.MeasIIoPinList);
                    break;
                case "Z":
                    arrstr[1] = PinConvertMapping(arrstr[1].ToUpper());
                    nHardIpInfo.MeasZStr = arrstr[1];
                    AddTokens(arrstr[1], nHardIpInfo.MeasZPinList);
                    break;
                case "R":
                    arrstr[1] = PinConvertMapping(arrstr[1].ToUpper());
                    nHardIpInfo.MeasRStr = arrstr[1];
                    AddTokens(arrstr[1], nHardIpInfo.MeasRPinList);
                    break;
                case "V":
                    arrstr[1] = PinConvertMapping(arrstr[1].ToUpper());
                    nHardIpInfo.MeasVStr = arrstr[1];
                    AddTokensWithVddContainsSplit(arrstr[1], nHardIpInfo.MeasVPowerPinList, nHardIpInfo.MeasVIoPinList);
                    break;
                case "VDIFF":
                    nHardIpInfo.MeasVdiffStr = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasVdiffPinList);
                    break;
                case "VDIFF2":
                    nHardIpInfo.MeasVdiff2Str = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasVdiff2PinList);
                    break;
                case "IDIFF":
                    nHardIpInfo.MeasIdiffStr = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasIdiffPinList);
                    break;
                case "FDIFF":
                    nHardIpInfo.MeasFdiffStr = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasFdiffPinList);
                    break;
                case "VOCM":
                    nHardIpInfo.MeasVocmStr = arrstr[1].ToUpper();
                    AddTokens(arrstr[1], nHardIpInfo.MeasVocmPinList);
                    break;
            }
        }

        private static void AddTokens(string str, List<string> target)
        {
            string[] tokens = Regex.Split(str.Replace("+", ","), ",").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            foreach (string token in tokens)
            {
                target.Add(token);
            }
        }

        private static void AddTokensWithVddContainsSplit(string str, List<string> powerTarget, List<string> ioTarget)
        {
            string[] tokens = Regex.Split(str.Replace("+", ","), ",").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            foreach (string token in tokens)
            {
                if (token.Contains("VDD", StringComparison.OrdinalIgnoreCase))
                {
                    powerTarget.Add(token);
                }
                else
                {
                    ioTarget.Add(token);
                }
            }
        }

        private static List<HardIpSeqInfo> ExtractMeasSeqInfo(HardIpReference info)
        {
            var mMeasSeqInfo = new List<HardIpSeqInfo>();
            // MeasSeqList = RemoveDummy(MeasSeqStr).Split(',');
            //           MeasSeqList = Regex.Split(RemoveDummy(MeasSeqStr), ",").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            if (RemoveDummy(info.MeasSeqStr) == "")
            {
                return mMeasSeqInfo;
            }

            string[] measSeqList = Regex.Split(RemoveDummy(info.MeasSeqStr), ",").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            string[] measFList = SplitPlusOrNull(info.MeasFStr);
            string[] measIList = SplitPlusOrNull(info.MeasIStr);
            string[] measVList = SplitPlusOrNull(info.MeasVStr);
            string[] measVdiffList = SplitPlusOrNull(info.MeasVdiffStr);
            string[] measVdiff2List = SplitPlusOrNull(info.MeasVdiff2Str);
            string[] measIdiffList = SplitPlusOrNull(info.MeasIdiffStr);
            string[] measFdiffList = SplitPlusOrNull(info.MeasFdiffStr);
            string[] measVocmList = SplitPlusOrNull(info.MeasVocmStr);
            string[] measZList = SplitPlusOrNull(info.MeasZStr);
            string[] measRList = SplitPlusOrNull(info.MeasRStr);

            /*
                if (MeasVStr.IndexOf("+") != -1 || MeasIStr.IndexOf("+") != -1 || MeasFStr.IndexOf("+") != -1)
                {
                    Console.WriteLine("TP1");
                }*/

            for (int i = 0; i < measSeqList.Length; i++)
            {
                var nMeasSeqInfo = new HardIpSeqInfo { SeqName = measSeqList[i] };
                PopulateSeqInfo(nMeasSeqInfo, i, info, measFList, measIList, measVList, measVdiffList, measVdiff2List, measIdiffList, measFdiffList, measVocmList, measZList, measRList);
                mMeasSeqInfo.Add(nMeasSeqInfo);
            }
            return mMeasSeqInfo;
        }

        private static void PopulateSeqInfo(HardIpSeqInfo seq, int i, HardIpReference info,
            string[] measFList, string[] measIList, string[] measVList, string[] measVdiffList,
            string[] measVdiff2List, string[] measIdiffList, string[] measFdiffList,
            string[] measVocmList, string[] measZList, string[] measRList)
        {
            switch (seq.SeqName.ToUpper())
            {
                case "V":
                    seq.PinList = PickListEntry(measVList, i, info.MeasVStr);
                    AddSeqTokensVddRegexSplit(seq.PinList, seq.MeasVPowerPinList, seq.MeasVIoPinList);
                    break;
                case "I":
                    seq.PinList = PickListEntry(measIList, i, info.MeasIStr);
                    AddSeqTokensVddRegexSplit(seq.PinList, seq.MeasIPowerPinList, seq.MeasIIoPinList);
                    break;
                case "F":
                    seq.PinList = PickListEntry(measFList, i, info.MeasFStr);
                    AddSeqTokens(seq.PinList, seq.MeasFPinList);
                    break;
                case "VDIFF":
                    seq.PinList = PickListEntry(measVdiffList, i, info.MeasVdiffStr);
                    AddSeqTokens(seq.PinList, seq.MeasVdiffPinList);
                    break;
                case "VDIFF2":
                    seq.PinList = PickListEntry(measVdiff2List, i, info.MeasVdiff2Str);
                    AddSeqTokens(seq.PinList, seq.MeasVdiff2PinList);
                    break;
                case "IDIFF":
                    seq.PinList = PickListEntry(measIdiffList, i, info.MeasIdiffStr);
                    AddSeqTokens(seq.PinList, seq.MeasIdiffPinList);
                    break;
                case "FDIFF":
                    seq.PinList = PickListEntry(measFdiffList, i, info.MeasFdiffStr);
                    AddSeqTokens(seq.PinList, seq.MeasFdiffPinList);
                    break;
                case "VOCM":
                    seq.PinList = PickListEntry(measVocmList, i, info.MeasVocmStr);
                    AddSeqTokens(seq.PinList, seq.MeasVocmPinList);
                    break;
                case "Z":
                    seq.PinList = PickListEntry(measZList, i, info.MeasZStr);
                    AddSeqTokens(seq.PinList, seq.MeasZPinList);
                    break;
                case "R":
                    seq.PinList = PickListEntry(measRList, i, info.MeasRStr);
                    AddSeqTokens(seq.PinList, seq.MeasRPinList);
                    break;
            }
        }

        private static string[] SplitPlusOrNull(string s)
        {
            string cleaned = RemoveDummy(s);
            return cleaned == "" ? null : cleaned.Split('+');
        }

        private static string PickListEntry(string[] list, int i, string fallback)
        {
            return list.Length == 1 ? fallback : list[i];
        }

        private static void AddSeqTokens(string pinList, List<string> target)
        {
            string[] arrstr = Regex.Split(pinList, ",").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            foreach (string pin in arrstr)
            {
                target.Add(pin);
            }
        }

        private static void AddSeqTokensVddRegexSplit(string pinList, List<string> powerTarget, List<string> ioTarget)
        {
            string[] arrstr = Regex.Split(pinList, ",").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            foreach (string pin in arrstr)
            {
                if (Regex.IsMatch(pin, "VDD", RegexOptions.IgnoreCase))
                {
                    powerTarget.Add(pin);
                }
                else
                {
                    ioTarget.Add(pin);
                }
            }
        }

        private static string ExtractDigSrcInfo(string digSrcEqn)
        {
            string bit = "1";
            //wdr1_8+wdr2_8+wdr3_8+wdr4_8+wdr5_8+wdr6_8+wdr7_8+wdr8_8+wdr9_8+wdr10_8
            //wdr1_8+wdr2_8  wdr1_8=wr:0:7,wdr2_8=wr:0:7
            string[] arrstr = Regex.Split(RemoveDummy(digSrcEqn), "[+]").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            string[] arrstr1 = arrstr[0].Split('_');
            if (Regex.IsMatch(arrstr1[arrstr1.Length - 1], "[0-9]+"))
            {
                bit = arrstr1[arrstr1.Length - 1];
            }

            string assignMent = arrstr.Aggregate("", (current, step) => current + step + "=wr:0:" + (Convert.ToInt32(bit) - 1).ToString(CultureInfo.InvariantCulture) + ",");
            if (assignMent != "")
            {
                assignMent = assignMent.Substring(0, assignMent.Length - 1);
            }

            return bit + ";" + assignMent;
        }

        private static string RemoveDummy(string data)
        {
            string result = data;
            result = result.Replace(" ", "");
            result = result.Replace("\t", "");
            return result;
        }

        //key: origin in patinfo, value: replaced with "_"
        private string PinConvertMapping(string strIn)
        {
            string[] strInArr = strIn.Split('+', ',');
            string strOut = strIn.Replace("_", "");
            string[] strOutArr = strOut.Split('+', ',');

            for (int i = 0; i < strInArr.Length; i++)
            {
                string key = strInArr[i];
                if (key != "" && !PinUnderlineDict.ContainsKey(key))
                {
                    PinUnderlineDict.Add(key, strOutArr[i]);
                }
            }

            return strOut;
        }
    }
}
