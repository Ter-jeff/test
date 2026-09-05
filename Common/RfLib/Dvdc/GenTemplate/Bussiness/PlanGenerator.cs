using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Utility;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;
using RfLib.InstrumentSetup;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal partial class PlanGenerator(TemplateAutoGen templateAutoGen)
    {
        [GeneratedRegex("::")]
        private static partial Regex MyRegex6();
        [GeneratedRegex("::")]
        private static partial Regex MyRegex7();
        [GeneratedRegex(@"^(I|R1|R2)$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex8();
        [GeneratedRegex(@"^\[(?<setup>.*)\]$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex12();
        [GeneratedRegex("meas", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex13();
        [GeneratedRegex("pn|spur", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex14();
        [GeneratedRegex("::")]
        private static partial Regex MyRegex15();
        [GeneratedRegex("^UWS|^LPS|^LXS|^UPS", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex20();

        private readonly TemplateAutoGen _owner = templateAutoGen;

        private readonly Hashtable _describe = new()
        {
            {"F", "frequency"},
            {"V", "voltage"},
            {"I", "current"},
            {"R1", "Resistor"},
            {"R2", "Resistor"},
        };

        internal void GenPlanWithNonVdiff(List<TemplateRow> templateRows, HardIpSeqInfo hardIpSeqInfo, int blockindex, int stepIndex)
        {
            List<string> pinList = hardIpSeqInfo.PinList;
            pinList.Sort();
            List<string> forceList = [.. hardIpSeqInfo.ForceList.Split(',')];
            List<string> nameList = [.. hardIpSeqInfo.MeasName.Split(',')];
            bool singlePin = pinList.Count == 1;
            int pinIndex = 1;
            SortPinList(ref pinList, ref forceList, ref nameList);
            foreach (string pin in pinList)
            {
                //Meas information row
                var newTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
                templateRows.Add(newTempRow);

                string testName = UpdateTestName(nameList, pinIndex);
                if (!string.IsNullOrEmpty(testName))
                {
                    testName = string.Format("\"{0}\"", testName);
                }

                newTempRow.Meas = string.Format("Meas{0} Pin = {1} {2}", hardIpSeqInfo.SeqName, pin, testName);

                newTempRow.ForceCondition = UpdateForceCondition(hardIpSeqInfo, forceList, pin, pinIndex);
                newTempRow.Description = string.Format("Measure the {0} for {1}", _describe[hardIpSeqInfo.SeqName], pin);
                newTempRow.Description = UpdateDescription(hardIpSeqInfo, pin, newTempRow.ForceCondition, testName);
                newTempRow.Step = blockindex + "." + stepIndex;
                if (!singlePin)
                {
                    newTempRow.Step += "." + pinIndex;
                }

                pinIndex++;
            }
        }

        internal void GenPlanWithVdiff(List<TemplateRow> templateRows, HardIpSeqInfo hardIpSeqInfo, int blockindex, int stepIndex)
        {
            List<string> pinList = hardIpSeqInfo.PinList;
            List<string> forceList = [.. hardIpSeqInfo.ForceList.Split(',')];
            List<string> nameList = [.. hardIpSeqInfo.MeasName.Split(',')];

            int pinIndex = 1;
            char measType = hardIpSeqInfo.SeqName[0];
            var pdiffTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
            templateRows.Add(pdiffTempRow);
            //Measure voltage on PCIE_REF_CLK1_P. Force I = 0.8. Measure voltage on PCIE_REF_CLK2_P. Force I = 0.6.

            string pinPdiff = "";
            if (hardIpSeqInfo.PinList.Contains("::"))
            {
                pinPdiff = string.Join(",", pinList.Select(x => MyRegex6().Split(x)[0]));
            }

            int pNindex = 0;
            foreach (string pin in pinList.Select(x => MyRegex6().Split(x)[0]))
            {
                string forcestr = UpdateForceCondition(hardIpSeqInfo, forceList, pin, pNindex, "P");
                pdiffTempRow.ForceCondition += forcestr + ";";
                pdiffTempRow.Description += UpdateDescription(hardIpSeqInfo, pin, forcestr, "") + "\n";
                pNindex++;
            }

            pdiffTempRow.Meas = string.Format("Meas{0} Pin = {1}", measType, pinPdiff);
            pdiffTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            pinIndex++;

            var ndiffTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
            templateRows.Add(ndiffTempRow);
            string pinNdiff = "";
            if (hardIpSeqInfo.PinList.Contains("::"))
            {
                pinNdiff = string.Join(",", pinList.Select(x => MyRegex7().Split(x)[1]));
            }

            pNindex = 0;
            foreach (string pin in pinList.Select(x => MyRegex7().Split(x)[1]))
            {
                string forcestr = UpdateForceCondition(hardIpSeqInfo, forceList, pin, pNindex, "N");
                ndiffTempRow.ForceCondition += forcestr + ";";
                ndiffTempRow.Description += UpdateDescription(hardIpSeqInfo, pin, forcestr, "") + "\n";
                pNindex++;
            }
            ndiffTempRow.Meas = string.Format("Meas{0} Pin = {1}", measType, pinNdiff);
            ndiffTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            pinIndex++;

            var diffTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
            templateRows.Add(diffTempRow);

            diffTempRow.Description = string.Format("Calculate for {0}diff.", measType);
            if (nameList.Count != 0)
            {
                diffTempRow.Description += string.Format(" MeasureName for this {0}.", nameList[0]);
            }

            string pindiff = string.Join(",", hardIpSeqInfo.PinList);
            diffTempRow.Meas = string.Format("Meas{0}diff Pin = {1}", measType, pindiff);
            if (nameList.Count != 0)
            {
                diffTempRow.Meas += string.Format(" \"{0}\"", nameList[0]);
            }

            diffTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            pinIndex++;

            if (measType == 'V')
            {
                var vocmTempRow = new TemplateRow(blockindex, blockindex + "." + stepIndex);
                templateRows.Add(vocmTempRow);
                vocmTempRow.Description = "Calculate for Vocm.";
                vocmTempRow.Meas = "MeasVocm" + " Pin = " + pindiff;
                vocmTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
                pinIndex++;
            }
        }

        private static void SortPinList(ref List<string> pinlist, ref List<string> forceList, ref List<string> nameList)
        {

            bool isSingleForce = forceList.Count == 1;
            bool isSingleName = nameList.Count == 1;
            List<string> mergelist = [];
            int i = 0;
            //Merge pin with force condition, avoid same pin concern
            foreach (string pin in pinlist)
            {
                string forceVal = "";
                string nameVal = "";
                if (isSingleForce)
                {
                    forceVal = forceList[0];
                }
                else if (i < forceList.Count)
                {
                    forceVal = forceList[i];
                }
                else
                {
                    forceVal = "";
                }

                if (isSingleName)
                {
                    nameVal = nameList[0];
                }
                else if (i < nameList.Count)
                {
                    nameVal = nameList[i];
                }
                else
                {
                    nameVal = "";
                }

                mergelist.Add(string.Format("{0};{1};{2}", pin, forceVal, nameVal));
                i++;
            }
            pinlist.Sort();
            forceList = [];
            nameList = [];
            foreach (string pin in pinlist)
            {
                string item = mergelist.FirstOrDefault(p => p.Split(';')[0] == pin)!;
                forceList.Add(item.Split(';')[1]);
                nameList.Add(item.Split(';')[2]);
                mergelist.Remove(item);
            }
        }

        private static string UpdateTestName(List<string> nameList, int currentIndx)
        {
            string result = "";
            if (nameList.Count != 0)
            {
                result = nameList.Count >= currentIndx ? nameList[currentIndx - 1] : nameList[0];

            }
            return result.Trim('>');
        }

        private static string UpdateForceCondition(HardIpSeqInfo hardIpSeqInfo, List<string> forceList, string pin, int pinIndex, string diffType = "")
        {
            string result = "";
            if (diffType.Length == 0)
            {
                if (hardIpSeqInfo.SeqName.EqualsIgnoreCase("V"))
                {
                    if (forceList[pinIndex - 1].Length != 0)
                    {
                        result = string.Format("{0}:I:{1}", pin, forceList[pinIndex - 1]);
                    }
                }
                if (MyRegex8().IsMatch(hardIpSeqInfo.SeqName))
                {
                    if (forceList[pinIndex - 1].Length != 0)
                    {
                        result = string.Format("{0}:V:{1}", pin, forceList[pinIndex - 1]);
                    }
                }
            }
            else
            {
                string type = "";
                if (diffType == "P" && (pinIndex * 2) < forceList.Count)
                {
                    if (hardIpSeqInfo.SeqName.Contains('V'))
                    {
                        type = "I";
                    }
                    else if (hardIpSeqInfo.SeqName.Contains('I'))
                    {
                        type = "V";
                    }
                }
                else if (diffType == "N" && ((pinIndex * 2) + 1) < forceList.Count)
                {
                    if (hardIpSeqInfo.SeqName.Contains('V'))
                    {
                        type = "I";
                    }
                    else if (hardIpSeqInfo.SeqName.Contains('I'))
                    {
                        type = "V";
                    }
                }
                if (type.Length != 0)
                {
                    int index = diffType == "P" ? pinIndex * 2 : (pinIndex * 2) + 1;
                    result = string.Format("{0}:{1}:{2}", pin, type, forceList[index]);
                }

            }
            return result;
        }

        //Measure current on  DDRIOPINS. Force V = 0V.  MeasureName for this item is LOW.
        private string UpdateDescription(HardIpSeqInfo hardIpSeqInfo, string pin, string forcestr, string tnamestr)
        {
            string result;
            string forcePart = "";
            string tnamePart = "";

            if (forcestr.Length != 0)
            {
                string type = forcestr.Split(':')[1];
                string value = forcestr.Split(':')[2];
                string unit = type == "I" ? "A" : "V";

                forcePart = string.Format("Force {0} = {1}{2}", type, value, unit);
            }

            if (tnamestr.Length != 0)
            {
                tnamePart = "MeasureName for this item is " + tnamestr;
            }
            result = string.Format("Measure the {0} for {1}. {2} {3}", _describe[hardIpSeqInfo.SeqName], pin, forcePart, tnamePart);
            return result;
        }

        internal void GenPlanWithNonVdiffNew(string payload, List<TemplateRow> templateRows, HardIpSeqInfoNew hardIpSeqInfoNew, string seqindex, int blockindex, int stepIndex, bool isRFItem, bool isBBItem = false)
        {
            List<string> measpinList = hardIpSeqInfoNew.MeasPin;
            measpinList.Sort();

            List<string> nameList = hardIpSeqInfoNew.MeasName;
            bool singlePin = measpinList.Count == 1;
            int pinIndex = 1;
            int rfSubIdx = 0;

            foreach (MeasPin measPin in hardIpSeqInfoNew.MeasPins)
            {
                var testTypes = new List<string>();
                string path = "";
                if (!string.IsNullOrEmpty(measPin.RfInstrumentSetup))
                {
                    string info = MyRegex12().Match(measPin.RfInstrumentSetup).Groups["setup"].ToString();
                    foreach (string item in info.Split('$'))
                    {
                        if (!item.Contains('='))
                        {
                            continue;
                        }
                        string header = item.Split('=')[0].ToLower();
                        string data = item.Split('=')[1].ToLower();
                        if (header == "testtype")
                        {
                            testTypes = [.. data.Split(['+'], StringSplitOptions.RemoveEmptyEntries)];
                            //testTypes = data.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        }
                        if (header.EqualsIgnoreCase("Path"))
                        {
                            path = data;
                        }
                    }
                }
                if (isRFItem && (measPin.MeasType == MeasType.MeasV || measPin.MeasType == MeasType.MeasI))
                {
                    testTypes.Add(MyRegex13().Replace(measPin.MeasType, ""));
                }
                else if (testTypes.Count == 0)
                {
                    testTypes.Add("");
                }

                for (int i = 0; i < testTypes.Count; i++)
                {
                    //Meas information row
                    var newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex + i);
                    if (!isBBItem)
                    {
                        string testtype = testTypes[i];
                        int subCount = 1;
                        if (MyRegex14().IsMatch(testtype))
                        {
                            subCount = GetSubCountMeasWithSpecialTestType(testtype, measPin.RfInstrumentSetup);
                        }
                        for (int j = 0; j < subCount; j++)
                        {
                            newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex + i + "." + j);
                            templateRows.Add(newTempRow);
                            newTempRow.TestName = measPin.TestName;
                            newTempRow.PostCalc = testtype;
                            if (payload.Contains("oip3"))
                            {
                                ;
                            }
                            string instrumentSetup = GetInstrumentSetup(payload, measPin.RfInstrumentSetup, measPin.SequenceIndex, measPin.SubSequenceIndex);
                            if (!string.IsNullOrEmpty(instrumentSetup))
                            {
                                if (string.IsNullOrEmpty(newTempRow.PostCalc))
                                {
                                    newTempRow.PostCalc = SetDefaultTestType(instrumentSetup);
                                }
                                newTempRow.Meas = string.Format("{0} Pin = {1} ", measPin.MeasType, measPin.PinName);
                                newTempRow.InstrumentSetup = "RFInstSetup = " + instrumentSetup + ";";
                                if (!string.IsNullOrEmpty(path))
                                {
                                    newTempRow.InstrumentSetup += "Path = " + path + ";";
                                }
                            }

                            if (hardIpSeqInfoNew.ExpectValue.Count > 0 && !string.IsNullOrEmpty(hardIpSeqInfoNew.ExpectValue[0]))
                            {
                                newTempRow.Description += "ExpectValue: " + string.Join(",", hardIpSeqInfoNew.ExpectValue);
                            }
                            newTempRow.Step = blockindex + "." + stepIndex + i;
                            if (!singlePin)
                            {
                                newTempRow.Step += "." + pinIndex;
                            }
                            newTempRow.TestItem = blockindex;
                            newTempRow.Seqindex = seqindex;
                        }

                    }

                    newTempRow.ForceCondition = UpdateForceConditionNew(hardIpSeqInfoNew.ForceConditions);
                    string testName = UpdateTestName(nameList, pinIndex);
                    newTempRow.Meas = UpdateMeas(measPin, testName);

                    newTempRow.Description = UpdateDescriptionNew(hardIpSeqInfoNew.MeasSeq, measPin.PinName, newTempRow.ForceCondition, testName);
                    newTempRow.HiLimit.Add("CP1", measPin.HighLimit);
                    newTempRow.LoLimit.Add("CP1", measPin.LowLimit);
                    if (LocalSpecs.Options.Device == EnumDevice.LCD)
                    {
                        newTempRow.MiscInfo = UpdateLCDMiscInfo(hardIpSeqInfoNew, pinIndex - 1);
                    }
                    pinIndex++;
                }
                rfSubIdx++;
            }
        }

        private string GetInstrumentSetup(string payload, string instrumentSetup, int seqIndex, int subseqIndex)
        {
            InstrumentSetupRow? temp = _owner.InstrumentSetupForPatList.FirstOrDefault(x =>
                x.Pattern.EqualsIgnoreCase(payload) &&
                x.Pin.RfInstrumentSetup == instrumentSetup &&
                x.SubsetIndex == subseqIndex &&
                x.SeqIndex == seqIndex - 1);
            // 260318 add subsetting TBD_X_0 for draco
            InstrumentSetupRow? tBD_Subsetting = _owner.InstrumentSetupForPatList.FirstOrDefault(x => x.SubSetting == instrumentSetup);
            if (temp != null)
            {
                string rfSetup = temp.SetupName + "#" + temp.SubSetting;
                //InstrumentSetupForPatList.Remove(temp);
                return rfSetup;
            }
            else if (tBD_Subsetting != null) // 260318 add subsetting TBD_X_0 for draco
            {

                string rfSetup = tBD_Subsetting.SetupName + "#" + tBD_Subsetting.SubSetting;
                return rfSetup;
                ;
            }
            return "";
        }

        private static int GetSubCountMeasWithSpecialTestType(string type, string setupInfo)
        {
            Dictionary<string, string> specialMapDic = new Dictionary<string, string>(StringExtensions.IgnoreCase) { { "PN", "OffsetFreq" }, { "Spur", "rbw" } };
            int defaultCount = 1;
            if (string.IsNullOrEmpty(type))
            {
                return defaultCount;
            }
            if (!specialMapDic.TryGetValue(type, out string? dependentKey))
            {
                return defaultCount;
            }

            foreach (string setup in setupInfo.Split('$'))
            {
                if (!setup.Contains('='))
                {
                    continue;
                }
                if (setup.Split('=')[0].Trim().EqualsIgnoreCase(dependentKey))
                {
                    return setup.Split('=')[1].Split(';').Length;
                }
            }
            return defaultCount;
        }

        private string UpdateMeas(MeasPin measPin, string testName)
        {
            string measStr;
            if (measPin.MeasType == "N")
            {
                measStr = MeasType.MeasN;
            }
            else
            {
                //measStr = string.Format("{0} Pin = {1}", measPin.MeasType, measPin.PinName);
                measStr = string.Format("{0} Pin = {1}", measPin.MeasType, _owner.SearchPinInChannelMap(measPin));
            }

            if (!string.IsNullOrEmpty(measPin.CusStr))
            {
                measStr += string.Format(@"({0})", measPin.CusStr);
            }
            if (!string.IsNullOrEmpty(testName))
            {
                measStr += string.Format(" \"{0}\"", testName);
            }
            return measStr;
        }

        internal string UpdateForceConditionNew(ForceCondition forceCondition)
        {
            var result = new List<string>();
            foreach (ForcePin forcepin in forceCondition.ForcePins)
            {
                if (forcepin.ForceType.ContainsIgnoreCase("sweep"))
                {
                    foreach (string pin in forcepin.PinName.Split(','))
                    {
                        string sweepPin = pin;

                        result.Add(string.Format("{0}:{1}:{2}", pin, TemplateAutoGen.MyRegex11().Replace(forcepin.ForceType, ""), sweepPin));
                    }
                }
                else if (forcepin.ForceType.EqualsIgnoreCase("ac"))
                {
                    result.Add(string.Format("{1}:{0}:{2}", _owner.SearchPinInChannelMap(new MeasPin(forcepin.PinName, "")), forcepin.ForceType, forcepin.ForceValue));
                }
                else
                {
                    result.Add(string.Format("{0}:{1}:{2}", _owner.SearchPinInChannelMap(new MeasPin(forcepin.PinName, "")), forcepin.ForceType, forcepin.ForceValue));
                }
            }

            return string.Join(";\n", result);
        }

        private string UpdateDescriptionNew(string seq, string measpin, string forcestr, string tnamestr)
        {
            string result;
            string forcePart = "";
            string tnamePart = "";

            if (forcestr.Length != 0)
            {
                var forceitems = new List<string>();
                foreach (string forceitem in forcestr.Split(';'))
                {
                    string pin = forceitem.Split(':')[0];
                    string type = forceitem.Split(':')[1];
                    string value = forceitem.Split(':')[2];
                    string unit = type.EqualsIgnoreCase("V") ? "V" : "A";
                    forceitems.Add(string.Format("Force Pin {0} with {1} = {2}{3}", pin, type, value, unit));

                }
                forcePart = string.Join(".", forceitems);
            }

            if (tnamestr.Length != 0)
            {
                tnamePart = "MeasureName for this item is " + tnamestr;
            }
            result = string.Format("Measure the {0} for {1}. ForceCondition with:{2}. Tname:{3}", _describe[seq], measpin, forcePart, tnamePart);
            return result;
        }

        internal void GenPlanWithVdiffNew(List<TemplateRow> templateRows, HardIpSeqInfoNew hardIpSeqInfoNew, int blockindex, int stepIndex)
        {
            List<string> pinList = hardIpSeqInfoNew.MeasPin;
            List<string> forceList = hardIpSeqInfoNew.ForceValue;
            List<string> nameList = hardIpSeqInfoNew.MeasName;

            int pinIndex = 1;
            string measType = hardIpSeqInfoNew.MeasSeq;
            var pdiffTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex);
            templateRows.Add(pdiffTempRow);
            //Measure voltage on PCIE_REF_CLK1_P. Force I = 0.8. Measure voltage on PCIE_REF_CLK2_P. Force I = 0.6.

            string pinPdiff = "";
            if (string.Join(",", hardIpSeqInfoNew.MeasPin).Contains("::"))
            {
                pinPdiff = string.Join(",", pinList.Select(x => MyRegex15().Split(x)[0]));
            }
            pdiffTempRow.ForceCondition = UpdateForceConditionNew(hardIpSeqInfoNew.ForceConditions);
            pdiffTempRow.Description = UpdateDescriptionNew(hardIpSeqInfoNew.MeasSeq, pinPdiff, pdiffTempRow.ForceCondition, "");

            pdiffTempRow.Meas = string.Format("Meas{0} Pin = {1}", measType[0], pinPdiff);
            pdiffTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            pinIndex++;

            var ndiffTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex);
            templateRows.Add(ndiffTempRow);
            string pinNdiff = "";
            if (string.Join(",", hardIpSeqInfoNew.MeasPin).Contains("::"))
            {
                pinNdiff = string.Join(",", pinList.Select(x => MyRegex15().Split(x)[1]));
            }

            ndiffTempRow.ForceCondition = UpdateForceConditionNew(hardIpSeqInfoNew.ForceConditions);
            ndiffTempRow.Description = UpdateDescriptionNew(hardIpSeqInfoNew.MeasSeq, pinNdiff, ndiffTempRow.ForceCondition, "");
            ndiffTempRow.Meas = string.Format("Meas{0} Pin = {1}", measType[0], pinNdiff);
            ndiffTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            pinIndex++;

            var diffTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex);
            templateRows.Add(diffTempRow);
            diffTempRow.ForceCondition = UpdateForceConditionNew(hardIpSeqInfoNew.ForceConditions);
            diffTempRow.Description = string.Format("Calculate for {0}diff.", measType[0]);
            if (nameList.Count != 0)
            {
                diffTempRow.Description += string.Format(" MeasureName for this {0}.", nameList[0]);
            }
            string pindiff = string.Join(",", string.Join(",", hardIpSeqInfoNew.MeasPin));
            diffTempRow.Meas = string.Format("Meas{0}diff Pin = {1}", measType[0], pindiff);
            if (nameList.Count != 0)
            {
                diffTempRow.Meas += string.Format(" \"{0}\"", nameList[0]);
            }
            if (hardIpSeqInfoNew.MeasPins.Any(p => !string.IsNullOrEmpty(p.HighLimit)))
            {
                diffTempRow.HiLimit.Add("CP1", hardIpSeqInfoNew.MeasPins.FirstOrDefault(p => !string.IsNullOrEmpty(p.HighLimit))!.HighLimit);
            }
            if (hardIpSeqInfoNew.MeasPins.Any(p => !string.IsNullOrEmpty(p.LowLimit)))
            {
                diffTempRow.LoLimit.Add("CP1", hardIpSeqInfoNew.MeasPins.FirstOrDefault(p => !string.IsNullOrEmpty(p.LowLimit))!.LowLimit);
            }
            if (hardIpSeqInfoNew.MeasPins.Count > 1)
            {
                ;
            }
            diffTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            pinIndex++;

            //if (seqinfo.MeasSeq[0] == 'V')
            //{
            //    var vocmTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex);
            //    blockTemplates.Add(vocmTempRow);
            //    vocmTempRow.Description = "Calculate for Vocm.";
            //    vocmTempRow.Meas = "MeasVocm" + " Pin = " + pindiff;
            //    vocmTempRow.Step = blockindex + "." + stepIndex + "." + pinIndex;
            //    pinIndex++;
            //}
        }

        private static string SetDefaultTestType(string instrumentSetup)
        {
            string instrument = instrumentSetup.Split('#')[1];
            return MyRegex20().IsMatch(instrument) ? "freq" : "";
        }

        private static string UpdateLCDMiscInfo(HardIpSeqInfoNew hardIpSeqInfoNew, int idx)
        {
            try
            {
                List<string> trimSet = [];
                var trimSubSet = new List<string>();
                int i = 0;
                string measStr = "";
                int j = 0;

                try
                {
                    foreach (char pin in hardIpSeqInfoNew.MeasPin[idx])
                    {
                        MeasPin measpin = hardIpSeqInfoNew.MeasPins[idx];
                        string irange = measpin.HighLimit;
                        string vrange = measpin.HighLimit;
                        vrange = vrange.Length == 0 ? "1.8" : vrange;
                        irange = irange.Length == 0 ? "0.02" : irange;
                        string time = string.IsNullOrEmpty(measpin.MeasWaitTime) ? "0.01" : measpin.MeasWaitTime;
                        var seqList =
                            hardIpSeqInfoNew.MeasSeq.Split('>').Where(p => !p.EqualsIgnoreCase("MeasWait")).ToList();
                        switch (seqList[0].ToUpper())
                        {
                            case "F":
                                measStr = string.Format("MeasWaitTime:{0}", time);
                                break;
                            case "VDM":
                                measStr = string.Format("MeasVRange:{0};MeasWaitTime:{1}", vrange, time);
                                break;
                            case "IDIFF":
                                measStr = string.Format("MeasIRange:{0};MeasWaitTime:{1}", irange, time);
                                break;
                            case "VDIFF":
                                measStr = string.Format("MeasVRange:{0};MeasWaitTime:{1}", vrange, time);
                                break;
                            case "N":
                                measStr = !string.IsNullOrEmpty(measpin.MeasWaitTime) ?
                                    string.Format("MeasWaitTime:{0}", measpin.MeasWaitTime) : "";
                                break;
                            case "I":
                                measStr = string.Format("MeasIRange:{0};MeasWaitTime:{1}", irange, time);
                                break;
                            case "V":
                                measStr = string.Format("MeasVRange:{0};MeasWaitTime:{1}", vrange, time);
                                break;
                        }

                        trimSubSet.Add(measStr);
                        j++;
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
                //foreach (var calcItem in Seq.Calc)
                //{
                //    TrimSubSet.Add(string.Format("Calc:{0};", calcItem.CalcEqn));
                //}
                //template.TestName = _GenMeasName(info).Replace("MeasName:", "");
                //template.HiLimit = ConvertValueUnit(info.NewInfo.HLimit);
                //template.LoLimit = ConvertValueUnit(info.NewInfo.LLimit);
                //template.Meas = string.Format("Meas{0} Pin = {1}", info.NewInfo.MeasSeq, info.NewInfo.MeasPin);
                if (trimSubSet.Distinct().Count() == 1)
                {
                    trimSet.Add(trimSubSet[0]);
                }
                else
                {
                    trimSet.Add(string.Join(";", trimSubSet));
                }

                i++;
                return string.Join(";", trimSet);

            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                return "";
            }
        }
    }
}
