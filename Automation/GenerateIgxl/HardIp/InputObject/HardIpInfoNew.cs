using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using CommonLib.Extension;

using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpInfoNew
    {
        public List<HardIpSeqInfoNew> SeqInfo;
        public List<string> TrimLoopHLimit = new List<string>();
        public List<string> TrimLoopLLimit = new List<string>();
        public string MeasSeq { get; set; } = "";
        public string MeasPin { get; set; } = "";
        public string MeasName { get; set; } = "";
        public string CalcEquation { get; set; } = "";
        public string CalcStoreName { get; set; } = "";
        public string MeasRange { get; set; } = "";
        public string MeasStoreName { get; set; } = "";
        public string ForceType { get; set; } = "";
        public string ForcePin { get; set; } = "";
        public string ForceValue { get; set; } = "";
        public string ExpectValue { get; set; } = "";
        public string ExpectFreq { get; set; } = "";
        public string HLimit { get; set; } = "";
        public string LLimit { get; set; } = "";
        public string RfSetup { get; set; } = "";
        public string MeasWait { get; set; } = "";

        public void ExtractMeasSeqInfo()
        {
            SeqInfo = new List<HardIpSeqInfoNew>();
            int i = 0;

            if (RemoveDummy(MeasSeq) != "")
            {
                List<string> measSeqList = RemoveDummy(MeasSeq).Split('|').ToList();
                /*
     * New Seq: ,,,F,I,V,Vdiff
New Seq MeasPin: |||PAD_CLKREF_0|PAD_DCTPN|DAC_ANATESTP::DAC_ANATESTN,ANALOG_TEST_P,ANALOG_TEST_N|
New Seq MeasName: |||220M_OUT|IPP|VREF|VDIFF
New Seq ForceType: |||I|V|V,SweepV|
New Seq ForcePin: |||ANALOG_TEST_P|ANALOG_TEST_N|PAD_ASG_TXN_0,PAD_ASG_TXN_1,PAD_ASG_TXN_2,PAD_ASG_TXN_3,PAD_ASG_TXN_4,PAD_ASG_TXN_5,PAD_ASG_TXN_6,PAD_ASG_TXN_7|
New Seq SweepPin: |||||PAD_ASG_TXP_0,PAD_ASG_TXP_1,PAD_ASG_TXP_2,PAD_ASG_TXP_3,PAD_ASG_TXP_4,PAD_ASG_TXP_5,PAD_ASG_TXP_6,PAD_ASG_TXP_7|
New Seq ForceValue: |||0.000025|0.5|0.4|
New Seq SweepValue: |||||0.25,0.55,0.1|
New Seq ExpectValue: ||2.2GHz|0.000025|0.6|0.5|
New Seq Hlimit:|||||0.7|
New Seq Llimit:|||||0.5|
     */
                List<string> measPinList = RemoveDummy(MeasPin).Split('|').ToList();

                List<string> measNameList = RemoveDummy(MeasName).Split('|').ToList();
                List<string> forcePinList = RemoveDummy(ForcePin).Split('|').ToList();
                List<string> forceTypeList = RemoveDummy(ForceType).Split('|').ToList();
                List<string> forceValueList = RemoveDummy(ForceValue).Split('|').ToList();
                List<string> expectValueList = RemoveDummy(ExpectValue).Split('|').ToList();
                List<string> expectFreqList = RemoveDummy(ExpectFreq).Split('|').ToList();
                List<string> hLimitList = RemoveDummy(HLimit).Split('|').ToList();
                List<string> lLimitList = RemoveDummy(LLimit).Split('|').ToList();
                List<string> measWaitList = RemoveDummy(MeasWait).Split('|').ToList();
                List<string> rfSetupList = RemoveDummy(RfSetup).Split('|').ToList();
                List<string> measStoreNameList = RemoveDummy(MeasStoreName).Split('|').ToList();
                List<string> calcEqnList = RemoveDummy(CalcEquation).Split('|').ToList();
                List<string> calcStoreNameList = RemoveDummy(CalcStoreName).Split('|').ToList();
                foreach (string seq in measSeqList)
                {
                    var nMeasSeqInfo = new HardIpSeqInfoNew { MeasSeq = seq.Equals("") || seq.Equals("MeasWait", StringComparison.OrdinalIgnoreCase) ? "N" : seq };
                    PopulateSeqInfoFields(nMeasSeqInfo, i, measPinList, forcePinList, forceValueList, forceTypeList,
                        measWaitList, expectValueList, expectFreqList, hLimitList, lLimitList);

                    nMeasSeqInfo.MeasPins = ExtractMeasPins(seq, measPinList, measStoreNameList, measNameList, rfSetupList, measWaitList, hLimitList, lLimitList, i);
                    nMeasSeqInfo.ForceConditions = ExtractForcePreMeas(forceTypeList, forcePinList, forceValueList, i);
                    nMeasSeqInfo.Calc = ExtractCalcEquation(calcEqnList, calcStoreNameList, measNameList, hLimitList, lLimitList, i);

                    if (measNameList != null && i < measNameList.Count)
                    {
                        nMeasSeqInfo.MeasName = measNameList[i].Split(',').ToList();
                    }

                    SeqInfo.Add(nMeasSeqInfo);
                    i++;
                }
            }


        }

        private static void PopulateSeqInfoFields(HardIpSeqInfoNew nMeasSeqInfo, int i, List<string> measPinList,
            List<string> forcePinList, List<string> forceValueList, List<string> forceTypeList, List<string> measWaitList,
            List<string> expectValueList, List<string> expectFreqList, List<string> hLimitList, List<string> lLimitList)
        {
            //MeasPin
            if (measPinList.Count > 0)
            {
                nMeasSeqInfo.MeasPin = measPinList.Count == 1
                    ? measPinList[0].Split(',').ToList()
                    : measPinList[i].Split(',').ToList();
            }

            //ForcePin
            if (forcePinList.Count > 0)
            {
                nMeasSeqInfo.ForcePin = forcePinList.Count == 1
                    ? forcePinList[0].Split(',').ToList()
                    : forcePinList[i].Split(',').ToList();
            }

            //ForceValue
            if (forceValueList.Count > 0)
            {
                nMeasSeqInfo.ForceValue = forceValueList.Count == 1
                    ? forceValueList[0].Split(',').ToList()
                    : forceValueList[i].Split(',').ToList();
            }

            //ForceType
            if (forceTypeList.Count > 0)
            {
                nMeasSeqInfo.ForceType = forceTypeList.Count == 1
                    ? forceTypeList[0].Split(',').ToList()
                    : forceTypeList[i].Split(',').ToList();
            }

            //MeasWaitTime
            if (measWaitList.Count > 0)
            {
                nMeasSeqInfo.MeasWait = measWaitList.Count == 1
                    ? measWaitList[0].Split(',').ToList()
                    : measWaitList[i].Split(',').ToList();
            }

            //ExpectValue
            if (expectValueList.Count > 0)
            {
                nMeasSeqInfo.ExpectValue = expectValueList.Count == 1
                    ? expectValueList[0].Split(',').ToList()
                    : expectValueList[i].Split(',').ToList();
            }

            //ExpectFreq
            if (expectFreqList.Count > 0)
            {
                nMeasSeqInfo.ExpectFreq = expectFreqList.Count == 1
                    ? expectFreqList[0].Split(',').ToList()
                    : expectFreqList[i].Split(',').ToList();
            }

            //HighLimit
            if (hLimitList.Count > 0)
            {
                nMeasSeqInfo.Hlimit = hLimitList.Count == 1
                    ? hLimitList[0].Split(',').ToList()
                    : hLimitList[i].Split(',').ToList();
            }

            //LowLimit
            if (lLimitList.Count > 0)
            {
                nMeasSeqInfo.Llimit = lLimitList.Count == 1
                    ? lLimitList[0].Split(',').ToList()
                    : lLimitList[i].Split(',').ToList();
            }
        }

        private ForceCondition ExtractForcePreMeas(List<string> forceTypes, List<string> forcePins, List<string> forceValues, int index)
        {
            var forcecondition = new ForceCondition();
            try
            {
                List<string> forcePinList = new List<string>();
                List<string> forceTypeList = new List<string>();
                List<string> forceValueList = new List<string>();
                var subForcePins = new List<string>();
                var subForceValues = new List<string>();
                if (index < forcePins.Count && !string.IsNullOrEmpty(forcePins[index]))
                {
                    forcePinList = RemoveDummy(forcePins[index]).Split('>').ToList();
                }

                if (index < forceTypes.Count && !string.IsNullOrEmpty(forceTypes[index]))
                {
                    forceTypeList = RemoveDummy(forceTypes[index]).Split('>').ToList();
                }

                if (index < forceValues.Count && !string.IsNullOrEmpty(forceValues[index]))
                {
                    forceValueList = RemoveDummy(forceValues[index]).Split('>').ToList();
                }

                int i = 0;

                foreach (string forcetype in forceTypeList)
                {
                    int interSeq = 0;
                    subForcePins = forcePinList[i].Split(',').ToList();
                    if (i < forceValueList.Count && !string.IsNullOrEmpty(forceValueList[i]))
                    {
                        subForceValues = RemoveDummy(forceValueList[i]).Split(',').ToList();
                    }

                    try
                    {
                        if (forcetype.ContainsIgnoreCase("SWEEP"))
                        {
                            foreach (string forcepin in subForcePins)
                            {
                                var force = new ForcePin
                                {
                                    PinName = forcepin,
                                    ForceType = forcetype,
                                    ForceValue = string.Join(",", subForceValues).Replace("A", "").Replace("V", "")
                                };
                                forcecondition.ForcePins.Add(force);
                            }
                        }
                        else if (forcetype.ContainsIgnoreCase("VDIFF"))
                        {
                            foreach (string forcepin in subForcePins)
                            {
                                string forceDiffInfos = "";
                                if (forceValueList.Count > 0)
                                {
                                    if (interSeq >= subForceValues.Count)
                                    {
                                        forceDiffInfos = subForceValues[0];
                                    }
                                    else
                                    {
                                        forceDiffInfos = subForceValues[interSeq];
                                    }
                                }

                                foreach (string forceDiffInfo in forceDiffInfos.Split(';'))
                                {

                                    var force = new ForcePin
                                    {
                                        PinName = forcepin,
                                        ForceType = forceDiffInfo.Split('&')[0],
                                        ForceValue = forceDiffInfo.Split('&')[1].Replace("A", "").Replace("V", "")
                                    };

                                    forcecondition.ForcePins.Add(force);
                                }

                                interSeq++;
                            }
                        }
                        else
                        {
                            foreach (string forcepin in subForcePins)
                            {
                                var force = new ForcePin { PinName = forcepin, ForceType = forcetype };

                                if (forceValueList.Count > 0)
                                {
                                    if (interSeq >= subForceValues.Count)
                                    {
                                        force.ForceValue = subForceValues[0].Replace("A", "").Replace("V", "");
                                    }
                                    else
                                    {
                                        force.ForceValue = subForceValues[interSeq].Replace("A", "").Replace("V", "");
                                    }
                                }
                                forcecondition.ForcePins.Add(force);
                                interSeq++;
                            }
                        }
                        i++;
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }

            return forcecondition;
        }

        private List<MeasPin> ExtractMeasPins(string measTypes, List<string> measPins, List<string> measStoreNames, List<string> measNames,
            List<string> instrumentSetups, List<string> measWaitTimes,
            List<string> hiLimits, List<string> loLimits, int index)
        {
            var allMeasPins = new List<MeasPin>();
            if (measTypes.Equals("N", StringComparison.OrdinalIgnoreCase) || measTypes.Equals("MeasWait", StringComparison.OrdinalIgnoreCase))
            {
                var measNPin = new MeasPin
                {
                    MeasType = "N",
                    RfInstrumentSetup = "TBD_" + index + "_0"
                };
                if (index < measWaitTimes.Count && !string.IsNullOrEmpty(measWaitTimes[index]))
                {
                    measNPin.MeasWaitTime = measWaitTimes[index];
                }

                allMeasPins.Add(measNPin);
                return allMeasPins;
            }
            try
            {
                List<string> measSeqList = measTypes.Split('>').ToList();

                List<string> measPinList = new List<string>();
                List<string> measNameList = new List<string>();
                List<string> measStoreNameList = new List<string>();
                List<string> instrumentSetupList = new List<string>();
                List<string> measWaitTimeList = new List<string>();
                List<string> hLimitList = new List<string>();
                List<string> lLimitList = new List<string>();
                var subPins = new List<string>();
                var subMeasNames = new List<string>();
                var subMeasStoreNames = new List<string>();
                var subInstrumentSetup = new List<string>();
                var subHLimits = new List<string>();
                var subLLimits = new List<string>();


                measPinList = measPins[index].Split('>').ToList();
                measNameList = SplitFieldByIndex(measNames, index);
                measStoreNameList = SplitFieldByIndex(measStoreNames, index);
                instrumentSetupList = SplitFieldByIndex(instrumentSetups, index);

                if (index < measWaitTimes.Count && !string.IsNullOrEmpty(measWaitTimes[index]))
                {
                    measWaitTimeList = measWaitTimes[index].Split('>').Where(p => !string.IsNullOrEmpty(p)).ToList();
                }

                hLimitList = SplitFieldByIndex(hiLimits, index);
                lLimitList = SplitFieldByIndex(loLimits, index);

                string freqList = SearchSweepFreq(instrumentSetupList);
                foreach (string freq in freqList.Split(';'))
                {
                    if (freqList.Split(';').Length > 1)
                    {
                    }
                    int seqindex = 0;
                    MeasPin waitPin = null;
                    foreach (string seq in measSeqList)
                    {
                        if (seq.Equals("CALC", StringComparison.Ordinal))
                        {
                            continue;
                        }

                        if (seq.Equals("MeasWait", StringComparison.Ordinal))
                        {
                            if (measWaitTimeList.Count > 0)
                            {
                                waitPin = new MeasPin
                                {
                                    MeasWaitTime = measWaitTimeList[0]
                                };
                                measWaitTimeList.RemoveAt(0);
                            }

                            continue;
                        }
                        int interSeq = 0;
                        subPins = measPinList[seqindex].Split(',').ToList();
                        UpdateSubList(ref subMeasNames, measNameList, seqindex, ',');
                        UpdateSubList(ref subMeasStoreNames, measStoreNameList, seqindex, ',');
                        UpdateSubList(ref subInstrumentSetup, instrumentSetupList, seqindex, '>');
                        UpdateSubList(ref subHLimits, hLimitList, seqindex, ',');
                        UpdateSubList(ref subLLimits, lLimitList, seqindex, ',');

                        foreach (string subpin in subPins)
                        {
                            try
                            {
                                var pin = new MeasPin
                                {
                                    MeasType = GetMeasType(seq)
                                };
                                if (waitPin != null)
                                {
                                    pin.MeasWaitTime = waitPin.MeasWaitTime;
                                    waitPin = null;
                                }
                                pin.SequenceIndex = index + 1;
                                pin.PinName = subpin;
                                //pin.PinName = DataConvertor.ConvertToNetName(subpin);
                                PopulateMeasPinDetails(pin, interSeq, freq, measNameList, subMeasNames, subMeasStoreNames, subInstrumentSetup, subHLimits, subLLimits);
                                allMeasPins.Add(pin);
                                interSeq++;
                            }
                            catch (Exception ex)
                            {
                                ErrorMessageBox.Show(string.Format(ex.ToString()));
                            }
                        }
                        seqindex++;
                    }
                    if (waitPin != null)
                    {
                        allMeasPins.ForEach(p => p.MeasWaitTime = p.MeasWaitTime + "&" + waitPin.MeasWaitTime);
                        waitPin = null;
                    }
                }

            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return allMeasPins;
        }

        private static List<string> SplitFieldByIndex(List<string> source, int index)
        {
            if (index < source.Count && !string.IsNullOrEmpty(source[index]))
            {
                return source[index].Split('>').ToList();
            }

            return new List<string>();
        }

        private static void UpdateSubList(ref List<string> sub, List<string> source, int seqindex, char separator)
        {
            if (seqindex < source.Count)
            {
                sub = source[seqindex].Split(separator).ToList();
            }
        }

        private void PopulateMeasPinDetails(MeasPin pin, int interSeq, string freq, List<string> measNameList,
            List<string> subMeasNames, List<string> subMeasStoreNames, List<string> subInstrumentSetup,
            List<string> subHLimits, List<string> subLLimits)
        {
            if (measNameList.Count > 0)
            {
                if (interSeq >= subMeasNames.Count)
                {
                    pin.TestName = subMeasNames[0];
                }
                else
                {
                    pin.TestName = subMeasNames[interSeq];
                }
            }
            if (subMeasStoreNames.Count > 0)
            {
                if (interSeq >= subMeasStoreNames.Count)
                {
                    pin.CusStr = subMeasStoreNames[0];
                }
                else
                {
                    pin.CusStr = subMeasStoreNames[interSeq];
                }
            }
            if (subInstrumentSetup.Count > 0)
            {
                if (interSeq >= subInstrumentSetup.Count)
                {
                    pin.RfInstrumentSetup = string.IsNullOrEmpty(freq) ? subInstrumentSetup[0] : ChangeSweepFreqContent(subInstrumentSetup[0], freq);
                }
                else
                {
                    pin.RfInstrumentSetup = string.IsNullOrEmpty(freq) ? subInstrumentSetup[interSeq] : ChangeSweepFreqContent(subInstrumentSetup[0], freq);
                }
            }

            if (subHLimits.Count > 0)
            {
                if (interSeq >= subHLimits.Count)
                {
                    pin.HighLimit = subHLimits[0];
                }
                else
                {
                    pin.HighLimit = subHLimits[interSeq];
                }
            }
            if (subLLimits.Count > 0)
            {
                if (interSeq >= subLLimits.Count)
                {
                    pin.LowLimit = subLLimits[0];
                }
                else
                {
                    pin.LowLimit = subLLimits[interSeq];
                }
            }
        }

        private string SearchSweepFreq(List<string> instrumentSetupList)
        {
            string result = "";
            int i = 0;
            string regInstr = @"\[(?<freq>.*)\]";
            foreach (string instrumentSetup in instrumentSetupList)
            {
                bool isFoundFreq = false;
                string inst = Regex.Match(instrumentSetup, regInstr, RegexOptions.IgnoreCase).Groups["freq"].Value;
                foreach (string item in inst.Split('$'))
                {
                    if (item.Split('=')[0].Trim().Equals("freq", StringComparison.OrdinalIgnoreCase))
                    {
                        isFoundFreq = true;
                        if (i == 0)
                        {
                            result = item.Split('=')[1].Trim();
                        }
                        else if (result != item.Split('=')[1].Trim())
                        {
                            return "";
                        }
                    }
                }
                i++;
                if (!isFoundFreq)
                {
                    return "";
                }
            }
            return result;
        }

        private string ChangeSweepFreqContent(string instrumentSetupInfo, string freqValue)
        {
            var result = new List<string>();
            string regInstr = @"\[(?<freq>.*)\]";
            if (Regex.IsMatch(instrumentSetupInfo, regInstr, RegexOptions.IgnoreCase))
            {
                string sweepFreq = Regex.Match(instrumentSetupInfo, regInstr, RegexOptions.IgnoreCase).Groups["freq"].Value;
                foreach (string item in sweepFreq.Split('$'))
                {
                    if (item.Split('=')[0].Trim().Equals("freq", StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add($"freq={freqValue}");
                    }
                    else
                    {
                        result.Add(item);
                    }
                }
                return "[" + string.Join("$", result) + "]";
            }
            return string.Join("$", result);
        }

        private List<MeasPin> ExtractCalcEquation(List<string> calcEqns, List<string> calcStoreNames, List<string> measNames, List<string> hiLimits, List<string> loLimits, int index)
        {
            var allCalcEqns = new List<MeasPin>();
            try
            {

                List<string> calcEqnList = new List<string>();
                List<string> calcStoreNameList = new List<string>();
                List<string> measNameList = new List<string>();
                List<string> hLimitList = new List<string>();
                List<string> lLimitList = new List<string>();

                if (index < calcEqns.Count && !string.IsNullOrEmpty(calcEqns[index]))
                {
                    calcEqnList = calcEqns[index].Split('>').ToList();
                }
                else
                {
                    return allCalcEqns;
                }

                if (index < calcStoreNames.Count && !string.IsNullOrEmpty(calcStoreNames[index]))
                {
                    calcStoreNameList = calcStoreNames[index].Split('>').ToList();
                }

                if (index < measNames.Count && !string.IsNullOrEmpty(measNames[index]))
                {
                    measNameList = measNames[index].Split('>').ToList();
                }

                if (index < hiLimits.Count && !string.IsNullOrEmpty(hiLimits[index]))
                {
                    hLimitList = hiLimits[index].Split('>').ToList();
                }

                if (index < loLimits.Count && !string.IsNullOrEmpty(loLimits[index]))
                {
                    lLimitList = loLimits[index].Split('>').ToList();
                }

                int seqindex = 0;

                foreach (string eqn in calcEqnList)
                {
                    if (string.IsNullOrEmpty(eqn))
                    {
                        seqindex++;
                        continue;
                    }
                    var calcItem = new MeasPin
                    {
                        MeasType = MeasType.MeasCalc,
                        TestName = measNameList[seqindex],
                        LowLimit = lLimitList[seqindex],
                        HighLimit = hLimitList[seqindex]
                    };
                    int eqnindex = 0;
                    foreach (string item in eqn.Split('>'))
                    {
                        switch (eqnindex)
                        {
                            case 0:
                                calcItem.CalcEqn = item;
                                break;
                            case 1:
                                calcItem.LowLimit = item;
                                break;
                            case 2:
                                calcItem.HighLimit = item;
                                break;
                        }
                        eqnindex++;
                    }
                    if (seqindex < calcStoreNameList.Count)
                    {
                        calcItem.CusStr = calcStoreNameList[seqindex];
                    }

                    allCalcEqns.Add(calcItem);
                    seqindex++;
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return allCalcEqns;
        }

        private string RemoveDummy(string data)
        {
            string result = data;
            result = result.Replace(" ", "");
            result = result.Replace("\t", "");
            return result;
        }

        private string GetMeasType(string seq)
        {
            switch (seq.ToUpper())
            {
                case "WISRC":
                case "WIMEAS":
                    return seq;
                case "MEASWAIT":
                    return "N";
                default:
                    return "Meas" + Regex.Replace(seq, "Meas", "", RegexOptions.IgnoreCase);
            }
        }
    }

    public class HardIpSeqInfo
    {
        public string SeqName { get; set; } = "";

        public List<string> PinList = new List<string>();

        public string ForceList = "";

        public List<string> ForceVPinList = new List<string>();
        public List<string> ForceIPinList = new List<string>();
        public List<MeasPin> MeasPins = new List<MeasPin>();
        public string MeasName = "";

        public HardIpSeqInfo Copy(HardIpSeqInfo seqInfo)
        {
            var newInfo = new HardIpSeqInfo
            {
                SeqName = seqInfo.SeqName,
                PinList = seqInfo.PinList,
                ForceList = seqInfo.ForceList,
                MeasName = seqInfo.MeasName,
                ForceVPinList = seqInfo.ForceVPinList,
                ForceIPinList = seqInfo.ForceIPinList
            };

            foreach (MeasPin pin in seqInfo.MeasPins)
            {
                var measPin = new MeasPin();
                measPin.Copy(pin);
                measPin.PinName = pin.PinName;
                measPin.SequenceIndex = pin.SequenceIndex;
                newInfo.MeasPins.Add(measPin);
            }

            return newInfo;
        }
    }

    public class PowerOnFly
    {
        public string SwitchSeq = "";
        public List<string> VMain = new List<string>();
        public List<string> VAlt = new List<string>();
    }

    public class HardIpSeqInfoNew
    {
        #region new pattern comment

        public string MeasSeq = "";
        public List<string> MeasPin = new List<string>();
        public List<string> MeasName = new List<string>();
        public List<string> ForceType = new List<string>();
        public List<string> ForcePin = new List<string>();
        public List<string> ForceValue = new List<string>();
        public List<string> ExpectValue = new List<string>();
        public List<string> ExpectFreq = new List<string>();
        public List<string> Hlimit = new List<string>();
        public List<string> Llimit = new List<string>();
        public List<string> MeasWait = new List<string>();
        public List<MeasPin> MeasPins = new List<MeasPin>();
        public ForceCondition ForceConditions = new ForceCondition();
        public List<MeasPin> Calc = new List<MeasPin>();
        #endregion
    }

}
