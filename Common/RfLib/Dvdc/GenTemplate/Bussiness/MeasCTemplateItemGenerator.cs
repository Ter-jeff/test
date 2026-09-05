using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;

using LogLib.Utility;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;
using RfLib.Dvdc.Reader.CapturePostProcess;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal static class MeasCTemplateItemGenerator
    {
        internal static List<MeasPin> GetSpecialLimits(HardIpInfo hardIpInfo)
        {
            var limitsList = new List<MeasPin>();
            //Limits:SNR,HLimit=0.5,Llimit=0.1
            foreach (string item in hardIpInfo.SpecialLimits)
            {
                var limit = new MeasPin
                {
                    MeasType = MeasType.MeasLimit
                };
                foreach (string category in item.Replace(":", "=").Split(';'))
                {
                    if (category.Contains('='))
                    {
                        string key = category.Split('=')[0];
                        string value = category.Split('=')[1];
                        switch (key.ToLower())
                        {
                            case "limits":
                                limit.TestName = value;
                                break;
                            case "highlimit":
                                limit.HighLimit = value;
                                break;
                            case "lowlimit":
                                limit.LowLimit = value;
                                break;
                            case "seqnum":
                                limit.SequenceIndex = int.Parse(value);
                                break;
                        }
                    }
                }
                limitsList.Add(limit);
            }
            return limitsList;
        }

        internal static void GenExtraLimit(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, int blockindex)
        {
            List<MeasPin> limits = GetSpecialLimits(hardIpInfo);

            foreach (MeasPin limit in limits)
            {
                var row = new WirelessTemplateRow(blockindex, "")
                {
                    Description = string.Format("Put Limits, TName:{0},HiLimit:{1},LoLimit:{2}", limit.TestName, limit.HighLimit, limit.LowLimit),
                    Meas = string.Format("Limits \"{0}\"", limit.TestName)
                };
                row.LoLimit.Add("CP1", limit.LowLimit);
                row.HiLimit.Add("CP1", limit.HighLimit);
                templateRows.Add(row);

            }
        }

        internal static void GenMeasCTemplateItem(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, int blockindex, int stepIndex)
        {

            string capPin = "JTAG_TDO";
            string[] capPinSet = hardIpInfo.CapPinName.Split('+');
            string[] capBitSet = hardIpInfo.CapBitName.Split('+');
            string[] capSigSet = hardIpInfo.CapBitStr.Split('+');
            string capBitName = "N/A";

            int subCount = 0;

            var sgmtBitsSet = capSigSet.Select(capSig => Convert.ToInt16(capSig.Split('_')[1])).Select(bit => (int)bit).ToList();
            var sgmtsSet = capSigSet.Select(capSig => capSig.Split('_')[0]).Select(sgmt => sgmt).ToList();
            var countBitsSet = sgmtBitsSet.Select(sgmtBit => 0).ToList();
            foreach (string capBit in hardIpInfo.CapBitStr.Split('+'))
            {

                if (capPinSet.Length > subCount)
                {
                    capPin = capPinSet[subCount];
                }

                if (capBitSet.Length > subCount)
                {
                    capBitName = capBitSet[subCount];
                }
                //MeasC information row
                var newTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
                templateRows.Add(newTempRow);
                string capbit = capBit.Split('_')[1];
                if (capBitName.Length != 0)
                {
                    newTempRow.Description = string.Format("Capture Code from {0} ({1}):{2}", capPin, capBitName, capbit);
                    newTempRow.Meas = string.Format(@"MeasC Pin = {0}:{1} ""{2}""", capPin, capbit, capBitName);
                }
                else
                {
                    newTempRow.Description = string.Format("Capture Code from {0} :{1}", capPin, capbit);
                    newTempRow.Meas = string.Format("MeasC Pin = {0}:{1}", capPin, capbit);
                }
                newTempRow.Step = blockindex + "." + stepIndex + "." + (subCount + 1);
                subCount++;
            }
        }

        internal static void GenMeasCTemplateItemNew(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, int blockindex, int stepIndex)
        {

            try
            {
                if (string.IsNullOrEmpty(hardIpInfo.CapPinName))
                {
                    return;
                }
                string capPin = "JTAG_TDO";
                string[] capPinSet = hardIpInfo.CapPinName.Split('+');
                string[] capBitSet = hardIpInfo.CapBitName.Split('+');
                string[] capSigSet = hardIpInfo.CapBitStr.Split('+');
                string capBitName = "N/A";

                int subCount = 0;

                var sgmtBitsSet =
                    capSigSet.Select(capSig => Convert.ToInt32(capSig.Split('_')[1])).Select(bit => bit).ToList();
                var sgmtsSet = capSigSet.Select(capSig => capSig.Split('_')[0]).Select(sgmt => sgmt).ToList();
                var countBitsSet = sgmtBitsSet.Select(sgmtBit => 0).ToList();
                List<string> storeName = string.IsNullOrEmpty(hardIpInfo.CapStoreName)
                    ? []
                    : [.. hardIpInfo.CapStoreName.Split(';')];
                var storeDic = new Dictionary<string, List<string>>();
                var usedStoreDic = new Dictionary<string, string>();
                storeDic = storeName.GroupBy(p => p.Split(':')[0]).ToDictionary(p => p.Key.ToLower(), p => p.ToList());
                if (storeName.Count > 1)
                {
                    ;
                }
                #region
                //if (TemplateGeneratorMain.IsSplitDssc)
                //{
                //    try
                //    {
                //        var dsscOutSet = patInfo.DsscOut.Replace("DSSC_OUT,", "").Split(',');
                //        foreach (var item in dsscOutSet)
                //        {
                //            var isSpilt = false;
                //            if (item.Split(':').Count() < 2)
                //                continue;
                //            //var DSSCcapBitName = item.Split(':')[1];
                //            var DSSCcapBitName = Regex.Match(item, @":(?<value>.*)").Groups["value"].Value;
                //            var capBit = item.Split(':')[0];
                //            var intcapBit = Convert.ToInt16(capBit);
                //            if (sgmtBitsSet[sgmtIndex] == intcapBit)
                //            {
                //                countBitsSet[sgmtIndex] = intcapBit;
                //            }
                //            else
                //            {
                //                isSpilt = true;
                //                countBitsSet[sgmtIndex] += intcapBit;
                //            }

                //            if (capPinSet.Count() == capSigSet.Count())
                //                capPin = capPinSet[sgmtIndex];
                //            else
                //                capPin = capPinSet[0];
                //            var newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex);
                //            blockTemplates.Add(newTempRow);
                //            var splitstr = "";
                //            if ((countBitsSet[sgmtIndex] != sgmtBitsSet[sgmtIndex]) || isSpilt)
                //            {
                //                splitstr = " ,Item Splited from:" + sgmtsSet[sgmtIndex];
                //            }
                //            if (capBitName != "")
                //            {
                //                newTempRow.Description = string.Format("Capture Code from {0}:{2}{3}", capPin,
                //                    DSSCcapBitName,
                //                    capBit, splitstr);
                //                if (StoreDic.ContainsKey(DSSCcapBitName) && StoreDic[DSSCcapBitName].Count > 0)
                //                {
                //                    newTempRow.Meas = string.Format(@"MeasC Pin = {0}({3}):{2} ""{1}""", capPin, DSSCcapBitName,
                //                        capBit, StoreDic[DSSCcapBitName][0].Split(':')[1]);
                //                    if (!usedStoreDic.ContainsKey(DSSCcapBitName))
                //                        usedStoreDic.Add(DSSCcapBitName, StoreDic[DSSCcapBitName][0]);
                //                    StoreDic[DSSCcapBitName].RemoveAt(0);
                //                }
                //                else if (usedStoreDic.ContainsKey(DSSCcapBitName))
                //                {
                //                    newTempRow.Meas = string.Format(@"MeasC Pin = {0}({3}):{2} ""{1}""", capPin, DSSCcapBitName,
                //                        capBit, StoreDic[DSSCcapBitName][0].Split(':')[1]);
                //                }
                //                else
                //                    newTempRow.Meas = string.Format(@"MeasC Pin = {0}:{2} ""{1}""", capPin,
                //                        DSSCcapBitName,
                //                        capBit
                //                        );
                //            }
                //            else
                //            {
                //                newTempRow.Description = string.Format("Capture Code from {0} :{1}{2}", capPin,
                //                    capBit, splitstr);
                //                newTempRow.Meas = string.Format("MeasC Pin = {0}:{1}", capPin, capBit);
                //            }
                //            newTempRow.Step = blockindex + "." + stepIndex + "." + (subCount + 1);
                //            subCount++;

                //            if (sgmtBitsSet[sgmtIndex] == countBitsSet[sgmtIndex])
                //                sgmtIndex++;
                //        }
                //    }
                //    catch (Exception es)
                //    {
                //        throw es;
                //    }
                //}
                //else
                #endregion
                {
                    foreach (string capBit in hardIpInfo.CapBitStr.Split('+'))
                    {

                        if (capPinSet.Length > subCount)
                        {
                            capPin = capPinSet[subCount];
                        }
                        if (capBitSet.Length > subCount)
                        {
                            capBitName = capBitSet[subCount].ToLower();
                        }
                        //MeasC information row
                        var newTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
                        templateRows.Add(newTempRow);
                        string capbit = capBit.Split('_')[1];
                        if (capBitName.Length != 0)
                        {
                            newTempRow.Description = string.Format("Capture Code from {0}:{2}\"{1}\"", capPin, capBitName, capbit);
                            if (storeDic.TryGetValue(capBitName, out List<string>? value) && value.Count > 0)
                            {
                                newTempRow.Meas = string.Format(@"MeasC Pin = {0}({3}):{2} ""{1}""", capPin, capBitName, capbit, value[0].Split(':')[1]);
                                if (!usedStoreDic.ContainsKey(capBitName))
                                {
                                    usedStoreDic.Add(capBitName, storeDic[capBitName][0]);
                                }

                                value.RemoveAt(0);

                            }
                            else if (usedStoreDic.TryGetValue(capBitName, out string? value1))
                            {
                                newTempRow.Meas = string.Format(@"MeasC Pin = {0}({3}):{2} ""{1}""", capPin, capBitName, capbit, value1.Split(':')[1]);
                            }
                            else
                            {
                                newTempRow.Meas = string.Format(@"MeasC Pin = {0}:{1} ""{2}""", capPin, capbit, capBitName);
                            }
                        }
                        else
                        {
                            newTempRow.Description = string.Format("Capture Code from {0} :{1}", capPin, capbit);
                            newTempRow.Meas = string.Format("MeasC Pin = {0}:{1}", capPin, capbit);
                        }
                        newTempRow.Step = blockindex + "." + stepIndex + "." + (subCount + 1);
                        subCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        internal static void GenMeasCTemplateItemCpp(List<PostProcessSheetRow> postProcessSheetRows, List<TemplateRow> templateRows, int blockindex, int stepIndex)
        {

            try
            {
                string capPin = "JTAG_TDO";
                bool isCalcEquation = false;

                foreach (PostProcessSheetRow cpp in postProcessSheetRows)
                {
                    foreach (PostCalcInfo post in cpp.PostCalcs)
                    {
                        if (!string.IsNullOrEmpty(post.CalcEquation))
                        {
                            isCalcEquation = true;
                        }

                        if (!isCalcEquation || string.IsNullOrEmpty(post.CalcTestName))
                        {
                            continue;
                        }
                        var newTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
                        templateRows.Add(newTempRow);
                        newTempRow.Description = string.Format("Capture Code from {0}:{2}\"{1}\"", capPin, post.CalcTestName, post.Bit);

                        newTempRow.Meas = string.Format(@"MeasC Pin = {0}:{1} ""{2}""", capPin, post.Bit, post.CalcTestName);
                        //if(Regex.IsMatch(post.LowLimit,"NA",RegexOptions.IgnoreCase))
                        newTempRow.LoLimit.Add("CP1", post.LowLimit);
                        //if (Regex.IsMatch(post.HiLimit, "NA", RegexOptions.IgnoreCase))
                        newTempRow.HiLimit.Add("CP1", post.HiLimit);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }
    }
}
