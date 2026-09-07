using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

namespace Cautogen.AutoCZ.CharPostProcessor.InputReader
{
    internal class PatInfoReader
    {
        public static List<HardIpReference> Read(string hardIpRefFile)
        {
            GeneralFunc.WriteMessage("Reading HardIP info file " + hardIpRefFile);

            var mHardIpInfo = new List<HardIpReference>();
            var nHardIpInfo = new HardIpReference();
            int count = 0;
            bool start = false;
            int rowNum = 0;
            //string wrongLine = "";
            try
            {
                var nHardIpInfoFile = new StreamReader(hardIpRefFile);
                string param;
                while ((param = nHardIpInfoFile.ReadLine()) != null)
                {
                    rowNum++;
                    //wrongLine = param;
                    if (param.Contains("======"))
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
                    ParseLine(arrstr, nHardIpInfo);
                }
                if (mHardIpInfo.FindIndex(x => x.Payload.ToLower() == nHardIpInfo.Payload.ToLower()) == -1)
                {
                    nHardIpInfo.SeqInfo = ExtractMeasSeqInfo(nHardIpInfo);
                    mHardIpInfo.Add(nHardIpInfo);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Extracting patInfo file failed at row: " + rowNum + "  ---\r\n " + "\r\n" + ex.Message);
            }
            return mHardIpInfo;
        }

        private static void HandleSeparator(List<HardIpReference> mHardIpInfo, ref HardIpReference nHardIpInfo, ref int count, ref bool start, int rowNum)
        {
            if (count == 0)
            {
                start = true;
            }
            else
            {
                HardIpReference current = nHardIpInfo;
                if (mHardIpInfo.FindIndex(x => x.Payload.ToLower() == current.Payload.ToLower()) == -1)
                {
                    try
                    {
                        current.SeqInfo = ExtractMeasSeqInfo(current);
                        mHardIpInfo.Add(current);
                    }
                    catch (Exception e)
                    {
                        string outString = string.Format("Extracting patInfo file failed at row: " + (rowNum - 1) + "  ---\r\n " + "\r\n" + e.Message);
                        GeneralFunc.WriteMessage(outString);
                    }
                }
                else
                {
                    string errorMessage = "Error: " + current.Payload + " is Duplicated in PatInfoFile.txt";
                    GeneralFunc.WriteMessage(errorMessage);
                }
                nHardIpInfo = new HardIpReference();
            }
            count++;
        }

        private static void ParseLine(string[] arrstr, HardIpReference nHardIpInfo)
        {
            if (ParseMetaField(arrstr, nHardIpInfo))
            {
                return;
            }
            ParseDataField(arrstr, nHardIpInfo);
        }

        private static bool ParseMetaField(string[] arrstr, HardIpReference nHardIpInfo)
        {
            string key = arrstr[0].ToLower();
            if (key.Contains("test_inst"))
            {
                nHardIpInfo.Payload = arrstr[1].ToLower();
                return true;
            }
            if (key.Contains("tset"))
            {
                nHardIpInfo.TimeSet = arrstr[1];
                return true;
            }
            if (key == "subr")
            {
                nHardIpInfo.Subr = arrstr[1].Trim();
                return true;
            }
            if (key.Contains("vm_vector"))
            {
                nHardIpInfo.Vm = arrstr[1];
                return true;
            }
            if (key == "capbit")
            {
                nHardIpInfo.CapBit = Convert.ToInt32(arrstr[1]);
                return true;
            }
            if (key.Contains("capbitstr"))
            {
                nHardIpInfo.CapBitStr = arrstr[1].Trim();
                return true;
            }
            if (key.Contains("capbitname"))
            {
                nHardIpInfo.CapBitName = arrstr[1].Trim();
                return true;
            }
            if (key.Contains("sendbitname"))
            {
                nHardIpInfo.SendBitName = arrstr[1].Trim();
                return true;
            }
            if (key.Contains("cappinname"))
            {
                nHardIpInfo.CapPinName = arrstr[1].Trim();
                return true;
            }
            if (key.Contains("sendpinname"))
            {
                nHardIpInfo.SendPinName = arrstr[1].Trim();
                return true;
            }
            if (key == "sendbit")
            {
                nHardIpInfo.SendBit = Convert.ToInt32(arrstr[1]);
                return true;
            }
            if (key.Contains("dssc_out"))
            {
                nHardIpInfo.DsscOut = string.Join(":", arrstr);
                return true;
            }
            return false;
        }

        private static void ParseDataField(string[] arrstr, HardIpReference nHardIpInfo)
        {
            string key = arrstr[0].ToLower();
            if (key.Contains("sendbitstr"))
            {
                nHardIpInfo.SendBitStr = arrstr[1].Trim();
                if (RemoveDummy(nHardIpInfo.SendBitStr) != "")
                {
                    string tmpstr = ExtractDigSrcInfo(RemoveDummy(nHardIpInfo.SendBitStr));
                    string[] arrstr1 = tmpstr.Split(';');
                    nHardIpInfo.DigSrcDataWidth = arrstr1[0];
                    nHardIpInfo.DigSrcAssign = arrstr1[1];
                }
                return;
            }
            if (key.Contains("measseq"))
            {
                nHardIpInfo.MeasSeqStr = arrstr[1].ToUpper();
                return;
            }
            ParseMeasField(arrstr, nHardIpInfo);
        }

        private static void ParseMeasField(string[] arrstr, HardIpReference nHardIpInfo)
        {
            switch (arrstr[0].ToLower())
            {
                case "f":
                    arrstr[1] = arrstr[1].ToUpper().Replace("&", "::");
                    nHardIpInfo.MeasFStr = arrstr[1];
                    AddTokens(arrstr[1], nHardIpInfo.MeasFPinList);
                    break;
                case "i":
                    arrstr[1] = arrstr[1].ToUpper();
                    nHardIpInfo.MeasIStr = arrstr[1];
                    AddTokensWithVddSplit(arrstr[1], nHardIpInfo.MeasIPowerPinList, nHardIpInfo.MeasIioPinList);
                    break;
                case "r1":
                    arrstr[1] = arrstr[1].ToUpper();
                    nHardIpInfo.MeasR1Str = arrstr[1];
                    AddTokens(arrstr[1], nHardIpInfo.MeasR1PinList);
                    break;
                case "r2":
                    arrstr[1] = arrstr[1].ToUpper();
                    nHardIpInfo.MeasR2Str = arrstr[1];
                    AddTokens(arrstr[1], nHardIpInfo.MeasR2PinList);
                    break;
                case "v":
                    arrstr[1] = arrstr[1].ToUpper();
                    nHardIpInfo.MeasVStr = arrstr[1];
                    AddTokensWithVddSplit(arrstr[1], nHardIpInfo.MeasVPowerPinList, nHardIpInfo.MeasVioPinList);
                    break;
                case "vdiff":
                    //arrstr[1] = HardIpUtilityMain.DataConvertor.ConvertToNetName(arrstr[1].ToUpper());
                    nHardIpInfo.MeasVdiffStr = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasVdiffPinList);
                    break;
                case "vdiff2":
                    //arrstr[1] = HardIpUtilityMain.DataConvertor.ConvertToNetName(arrstr[1].ToUpper());
                    nHardIpInfo.MeasVdiff2Str = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasVdiff2PinList);
                    break;
                case "idiff":
                    //arrstr[1] = HardIpUtilityMain.DataConvertor.ConvertToNetName(arrstr[1].ToUpper());
                    nHardIpInfo.MeasIdiffStr = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasIdiffPinList);
                    break;
                case "fdiff":
                    //arrstr[1] = HardIpUtilityMain.DataConvertor.ConvertToNetName(arrstr[1].ToUpper());
                    nHardIpInfo.MeasFdiffStr = arrstr[1].ToUpper().Replace("&", "::");
                    AddTokens(arrstr[1], nHardIpInfo.MeasFdiffPinList);
                    break;
                case "vocm":
                    //arrstr[1] = HardIpUtilityMain.DataConvertor.ConvertToNetName(arrstr[1].ToUpper());
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

        private static void AddTokensWithVddSplit(string str, List<string> powerTarget, List<string> ioTarget)
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
            string[] measR1List = SplitPlusOrNull(info.MeasR1Str);
            string[] measR2List = SplitPlusOrNull(info.MeasR2Str);

            for (int i = 0; i < measSeqList.Length; i++)
            {
                var nMeasSeqInfo = new HardIpSeqInfo { SeqName = measSeqList[i] };
                PopulateSeqInfo(nMeasSeqInfo, i, info, measFList, measIList, measVList, measVdiffList, measVdiff2List, measIdiffList, measFdiffList, measVocmList, measR1List, measR2List);
                mMeasSeqInfo.Add(nMeasSeqInfo);
            }
            return mMeasSeqInfo;
        }

        private static void PopulateSeqInfo(HardIpSeqInfo seq, int i, HardIpReference info,
            string[] measFList, string[] measIList, string[] measVList, string[] measVdiffList,
            string[] measVdiff2List, string[] measIdiffList, string[] measFdiffList,
            string[] measVocmList, string[] measR1List, string[] measR2List)
        {
            string seqLower = seq.SeqName.ToLower();
            switch (seqLower)
            {
                case "v":
                    seq.PinList = PickListEntry(measVList, i, info.MeasVStr);
                    AddSeqTokensVddSplit(seq.PinList, seq.MeasVPowerPinList, seq.MeasVioPinList);
                    break;
                case "i":
                    seq.PinList = PickListEntry(measIList, i, info.MeasIStr);
                    AddSeqTokensVddSplit(seq.PinList, seq.MeasIPowerPinList, seq.MeasIioPinList);
                    break;
                case "f":
                    seq.PinList = PickListEntry(measFList, i, info.MeasFStr);
                    AddSeqTokens(seq.PinList, seq.MeasFPinList);
                    break;
                case "vdiff":
                    seq.PinList = PickListEntry(measVdiffList, i, info.MeasVdiffStr);
                    AddSeqTokens(seq.PinList, seq.MeasVdiffPinList);
                    break;
                case "vdiff2":
                    seq.PinList = PickListEntry(measVdiff2List, i, info.MeasVdiff2Str);
                    AddSeqTokens(seq.PinList, seq.MeasVdiff2PinList);
                    break;
                case "idiff":
                    seq.PinList = PickListEntry(measIdiffList, i, info.MeasIdiffStr);
                    AddSeqTokens(seq.PinList, seq.MeasIdiffPinList);
                    break;
                case "fdiff":
                    seq.PinList = PickListEntry(measFdiffList, i, info.MeasFdiffStr);
                    AddSeqTokens(seq.PinList, seq.MeasFdiffPinList);
                    break;
                case "vocm":
                    seq.PinList = PickListEntry(measVocmList, i, info.MeasVocmStr);
                    AddSeqTokens(seq.PinList, seq.MeasVocmPinList);
                    break;
                case "r1":
                    seq.PinList = PickListEntry(measR1List, i, info.MeasR1Str);
                    AddSeqTokens(seq.PinList, seq.MeasR1PinList);
                    break;
                case "r2":
                    seq.PinList = PickListEntry(measR2List, i, info.MeasR2Str);
                    AddSeqTokens(seq.PinList, seq.MeasR2PinList);
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

        private static void AddSeqTokensVddSplit(string pinList, List<string> powerTarget, List<string> ioTarget)
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
            // wdr1_8+wdr2_8+wdr3_8+wdr4_8+wdr5_8+wdr6_8+wdr7_8+wdr8_8+wdr9_8+wdr10_8
            // wdr1_8+wdr2_8  wdr1_8=wr:0:7,wdr2_8=wr:0:7
            string bit = "1";
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
    }
}
