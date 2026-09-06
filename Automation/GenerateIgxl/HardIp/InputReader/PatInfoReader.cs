using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using LogLib.Static;
using LogLib.Utility;

using TestPlanLib.PowerMerge;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class PatInfoReader
    {
        public Dictionary<string, List<string>> PinGroupList = new Dictionary<string, List<string>>();
        public PowerMerge PowerMerge;
        private static readonly Regex _regex = new Regex(":(?<value>.*)");

        public PatInfoReader()
        {
            if (EpWorkbook.TestPlanWorkbook != null && TestPlanStatic.PowerMergeSheet != null)
            {
                PowerMerge = TestPlanStatic.PowerMergeSheet.PowerMerge;
            }
            else
            {
                PowerMerge = new PowerMerge();
            }

            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap == null)
            {
                if (EpWorkbook.TestPlanWorkbook != null)
                {
                    if (EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinMap"] != null)
                    {
                        var pinmapReader = new ReadPinMapSheet();
                        pinMap = pinmapReader.ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinMap"]);
                    }
                    else
                    {
                        //const string outString = "Missing IO_PinMap sheet in TestPlan";
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_MissingNeededSheets_02,
                            "",
                            1,
                            0,
                            ["IO_PinMap"]
                        );
                        return;
                    }
                }

            }
            if (pinMap != null)
            {
                foreach (PinGroup pinGroup in pinMap.GroupList)
                {
                    if (!PinGroupList.ContainsKey(pinGroup.PinName.ToUpper()))
                    {
                        PinGroupList.Add(pinGroup.PinName.ToUpper(), pinGroup.PinList.Select(x => x.PinName).ToList());
                    }
                }
            }
        }

        public List<HardIpInfo> ExtractHardIpInfos(string hardIpRefFile)
        {
            string[] lines = File.ReadAllLines(hardIpRefFile);
            return ExtractHardIpInfos(lines);
        }

        public List<HardIpInfo> ExtractHardIpInfos(IEnumerable<string> lines)
        {
            var hardIpReferences = new List<HardIpInfo>();
            var existList = new HashSet<string>();
            var nHardIpInfo = new HardIpInfo();
            string[] arrstr = new string[] { };
            int count = 0;
            bool start = false;
            int index;
            int rowNum = 0;
            string rawStr;
            try
            {
                foreach (string line in lines)
                {
                    string param = line;
                    rowNum++;
                    rawStr = param;
                    if (param.IndexOf("======", StringComparison.Ordinal) != -1)
                    {
                        if (count == 0)
                        {
                            start = true;
                        }
                        else
                        {
                            if (!existList.Contains(nHardIpInfo.FullPattern.ToUpper()))
                            {
                                nHardIpInfo.PatInfoExist = true;
                                if (!nHardIpInfo.IllegalChar)
                                {
                                    try
                                    {
                                        nHardIpInfo.SeqInfo = ExtractMeasSeqInfoNew(nHardIpInfo.MeasSeqStr, nHardIpInfo.MeaSet, nHardIpInfo.MeasName);
                                        nHardIpInfo.NewInfo?.ExtractMeasSeqInfo();

                                        if (nHardIpInfo.NewInfo != null)
                                        {
                                            HardIpInfo tmpHardIpInfo = ConvertHardIpInfoNewToOld(nHardIpInfo, nHardIpInfo.NewInfo);
                                            List<HardIpSeqInfo> convertInfo = ExtractMeasSeqInfoNew(tmpHardIpInfo.MeasSeqStr, tmpHardIpInfo.MeaSet, tmpHardIpInfo.MeasName);
                                            nHardIpInfo.SeqInfo = convertInfo;
                                            hardIpReferences.Add(nHardIpInfo);

                                        }
                                        else
                                        {
                                            hardIpReferences.Add(nHardIpInfo);
                                        }

                                        existList.Add(nHardIpInfo.FullPattern.ToUpper());
                                    }
                                    catch (Exception e)
                                    {
                                        Response.Report(string.Format("Extracting patInfo file failed at row: " + (rowNum - 1) + "  ---\r\n " + "\r\n" + e.StackTrace), EnumMessageLevel.Error, 50);
                                    }
                                }
                            }
                            else
                            {
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_DuplicatePatternInPatInfo_01,
                                    string.Empty,
                                    0,
                                    0,
                                    [nHardIpInfo.Payload]
                                );
                            }
                            nHardIpInfo = new HardIpInfo();
                        }
                        count++;
                    }
                    if (start && !param.Contains("======"))
                    {
                        param = param.Replace(" ", "");
                        if (param.LastIndexOf(":-") != param.IndexOf(":-"))
                        {
                            param = param.Replace(":-", "&");
                            arrstr = param.Split('&');
                            param = arrstr[0] + ":-" + arrstr[1] + "&-" + arrstr[2];
                        }
                        else if (param.Split(':').Length >= 3)
                        {
                            param = param.Replace(":-", "&-");
                        }
                        param = param.Replace("::", "&");
                        arrstr = param.Split(':');
                        index = param.IndexOf(":", StringComparison.Ordinal);
                        string tmpstr;
                        if (index != -1)
                        {
                            tmpstr = param.Substring(index + 1, param.Length - index - 1);
                        }
                        else
                        {
                            tmpstr = param;
                        }
                        tmpstr = tmpstr.Replace("*", "");
                        if (tmpstr.Replace("*", "").Trim().Length == 0)
                        {
                            continue;
                        }

                        #region Analysis Information

                        string key = arrstr[0].Split('(')[0].ToLower().Trim();
                        string value = _regex.Match(param).Groups["value"].Value.Trim().Replace("&", "::");
                        ApplyKeyValue(nHardIpInfo, key, value, arrstr);

                        string text = arrstr[0].ToLower();
                        ApplyTextField(nHardIpInfo, text, arrstr, rawStr, tmpstr, param);
                        #endregion
                    }

                }
                if (hardIpReferences.FindIndex(x => string.Equals(x.FullPattern, nHardIpInfo.FullPattern, StringComparison.OrdinalIgnoreCase)) == -1)
                {
                    nHardIpInfo.PatInfoExist = true;
                    if (!nHardIpInfo.IllegalChar)
                    {
                        nHardIpInfo.SeqInfo = ExtractMeasSeqInfoNew(nHardIpInfo.MeasSeqStr, nHardIpInfo.MeaSet, nHardIpInfo.MeasName);
                        if (nHardIpInfo.NewInfo != null)
                        {
                            nHardIpInfo.NewInfo.ExtractMeasSeqInfo();
                            HardIpInfo tmpHardIpInfo = ConvertHardIpInfoNewToOld(nHardIpInfo, nHardIpInfo.NewInfo);
                            nHardIpInfo.SeqInfo = ExtractMeasSeqInfoNew(tmpHardIpInfo.MeasSeqStr, tmpHardIpInfo.MeaSet, tmpHardIpInfo.MeasName);
                            hardIpReferences.Add(nHardIpInfo);
                        }
                        else
                        {
                            hardIpReferences.Add(nHardIpInfo);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Extracting patInfo file failed at row: " + rowNum + "  ---\r\n " + "\r\n" + e.StackTrace);
            }
            return hardIpReferences.OrderByDescending(p => p.VmVector).ToList();
        }

        private static readonly HashSet<string> _measSetKeys = new HashSet<string>
        {
            "f", "v", "i", "r1", "r2", "vdiff", "vdiff2", "idiff", "fdiff", "vocm"
        };

        private void ApplyKeyValue(HardIpInfo nHardIpInfo, string key, string value, string[] arrstr)
        {
            if (_measSetKeys.Contains(key))
            {
                AddMeasSetGroup(nHardIpInfo, key, value);
                return;
            }

            switch (key)
            {
                case "time_domain":
                    nHardIpInfo.TimeDomain = value.ToLower();
                    break;
                case "test_inst":
                    nHardIpInfo.Payload = value.ToLower();
                    break;
                case "vm_vector":
                    nHardIpInfo.VmVector = value.ToLower();
                    break;
                case "tset":
                    nHardIpInfo.TimeSet = value.ToLower();
                    break;
                case "subr":
                    nHardIpInfo.Subroutine = value.ToLower();
                    break;
                case "call_subrs":
                    nHardIpInfo.CallSubrs = value.ToLower();
                    break;
                case "call_subrs_cnt":
                    nHardIpInfo.CallSubrsCnt = Convert.ToInt32(value);
                    break;
                case "call_subrs_cnt_diff":
                    nHardIpInfo.CallSubRoutinesCntDiff = value.ToLower();
                    break;
                #region digCap
                case "capbit":
                    nHardIpInfo.CapBit = Convert.ToInt32(value);
                    break;
                case "capbitstr":
                    nHardIpInfo.CapBitStr = value.ToLower();
                    break;
                case "capbitname":
                    nHardIpInfo.CapBitName = value.ToLower();
                    break;
                case "cappinname":
                    nHardIpInfo.CapPinName = value.ToLower();
                    break;
                #endregion
                default:
                    ApplyKeyValueExtra(nHardIpInfo, key, value, arrstr);
                    break;
            }
        }

        private void ApplyKeyValueExtra(HardIpInfo nHardIpInfo, string key, string value, string[] arrstr)
        {
            switch (key)
            {
                #region digSrc
                case "sendbit":
                    nHardIpInfo.SendBit = value;
                    break;
                case "sendbitstr":
                    nHardIpInfo.SendBitStr = value.ToLower();
                    break;
                case "sendbitname":
                    nHardIpInfo.SendBitName = value.ToLower();
                    break;
                case "sendpinname":
                    nHardIpInfo.SendPinName = value;
                    break;
                case "digsrcsignalname":
                    nHardIpInfo.DigSrcSignalName = value.ToLower();
                    break;
                #endregion
                case "xpins":
                    nHardIpInfo.Xpins = value;
                    break;
                case "unusediopins":
                    nHardIpInfo.UnusedIoPins = value;
                    break;
                case "measseq":
                    nHardIpInfo.MeasSeqStr = value.ToUpper();
                    break;
                case "measname":
                    nHardIpInfo.MeasName = arrstr[1].ToUpper();
                    break;
                case "forcev":
                    foreach (string pin in value.Split('+'))
                    {
                        nHardIpInfo.ForceVPinList.Add(pin);
                    }

                    break;
                case "forcei":
                    foreach (string pin in value.Split('+'))
                    {
                        nHardIpInfo.ForceIPinList.Add(pin);
                    }

                    break;
                case "version":
                    nHardIpInfo.Version = value.Replace("=", "").ToLower();
                    break;
                case "ssncorename":
                    nHardIpInfo.SsnCoreName = value.Replace("=", "").ToLower();
                    break;
                case "efuse_prog_time":
                    nHardIpInfo.EfuseRepeatCount = value.ToLower();
                    break;
                default:
                    break;
            }
        }

        private void AddMeasSetGroup(HardIpInfo nHardIpInfo, string key, string value)
        {
            nHardIpInfo.MeaSet.Add(key, new List<Dictionary<string, List<string>>>());
            foreach (string pins in value.Split('+'))
            {
                var seqPins = new Dictionary<string, List<string>>();
                foreach (string pin in pins.Split(','))
                {
                    if (!seqPins.ContainsKey(pin))
                    {
                        seqPins.Add(pin, new List<string>());
                    }

                    foreach (string pinGroupPin in DecompGroups(pin))
                    {
                        seqPins[pin].AddRange(ConvertToNetName(pinGroupPin).Split(','));
                    }
                }
                nHardIpInfo.MeaSet[key].Add(seqPins);
            }
        }

        private void ApplyTextField(HardIpInfo nHardIpInfo, string text, string[] arrstr, string rawStr, string tmpstr, string param)
        {
            if (ApplyNewSeqTextField(nHardIpInfo, text, arrstr, rawStr))
            {
                return;
            }

            ApplyTrimAndMiscTextField(nHardIpInfo, text, arrstr, tmpstr, param);
        }

        private bool ApplyNewSeqTextField(HardIpInfo nHardIpInfo, string text, string[] arrstr, string rawStr)
        {
            if (text.IndexOf("test_inst", StringComparison.Ordinal) != -1)
            {
                nHardIpInfo.Payload = arrstr[1].ToLower();
            }
            else if (text.Contains("dssc_out"))
            {
                nHardIpInfo.DsscOut = string.Join(":", arrstr).TrimEnd(',');
            }
            else if (text == "newseq")
            {
                nHardIpInfo.NewInfo = new HardIpInfoNew { MeasSeq = arrstr[1] };
            }
            else if (text == "newseqmeaspin")
            {
                nHardIpInfo.NewInfo.MeasPin = arrstr[1].ToUpper().Replace("&", "::");
            }
            else if (text == "newseqmeasname")
            {
                nHardIpInfo.NewInfo.MeasName = arrstr[1].ToUpper();
            }
            else if (text == "newseqforcetype")
            {
                nHardIpInfo.NewInfo.ForceType = arrstr[1];
            }
            else if (text == "newseqforcepin")
            {
                nHardIpInfo.NewInfo.ForcePin = arrstr[1];
            }
            else if (text == "newseqforcevalue")
            {
                nHardIpInfo.NewInfo.ForceValue = arrstr[1];
            }
            else if (text == "newseqdisconnectpins")
            {
                nHardIpInfo.DisConnectPins = arrstr[1];
            }
            else if (text == "newseqexpectvalue")
            {
                nHardIpInfo.NewInfo.ExpectValue = arrstr[1].ToUpper();
            }
            else if (text == "newseqhighlimit")
            {
                nHardIpInfo.NewInfo.HLimit = arrstr[1];
            }
            else if (text == "newseqlowlimit")
            {
                nHardIpInfo.NewInfo.LLimit = arrstr[1];
            }
            else if (text == "newseqtrimtarget")
            {
                nHardIpInfo.TrimTarget = arrstr[1];
            }
            else if (text == "newseqrfsetup")
            {
                List<string> rFinfo = rawStr.Split(':').ToList();
                rFinfo.RemoveAt(0);

                if (!string.IsNullOrEmpty(string.Join(":", rFinfo)))
                {
                    nHardIpInfo.NewInfo.RfSetup = string.Join(":", rFinfo);
                }
            }
            else if (text == "newseqmeaswait")
            {
                nHardIpInfo.NewInfo.MeasWait = arrstr[1];
            }
            else
            {
                return false;
            }

            return true;
        }

        private void ApplyTrimAndMiscTextField(HardIpInfo nHardIpInfo, string text, string[] arrstr, string tmpstr, string param)
        {
            if (text == "miscinfo")
            {
                List<string> miscinfo = arrstr.ToList();
                miscinfo.RemoveAt(0);
                nHardIpInfo.MiscInfo = new List<string> { string.Join(":", miscinfo) };
            }
            else if (text == "digsrcassignment")
            {
                List<string> data = arrstr.ToList();
                data.RemoveAt(0);
                nHardIpInfo.DigSrcAssignment = string.Join(":", data);
            }
            else if (text == "trimfusename")
            {
                nHardIpInfo.TrimFuseName = arrstr[1];
            }
            else if (text == "trimregname")
            {
                nHardIpInfo.TrimRegName.Add(arrstr[1]);
            }
            else if (text == "trimtype")
            {
                nHardIpInfo.TrimType = arrstr[1];
            }
            else if (text == "triminstname")
            {
                nHardIpInfo.TestName = arrstr[1];
            }
            else if (text == "bestcodecalcfunc")
            {
                nHardIpInfo.BestCodeCalcFunc = arrstr[1];
            }
            else if (text == "prepatforce" || text == "preforce")
            {
                List<string> force = arrstr.ToList();
                force.RemoveAt(0);
                nHardIpInfo.ForcePrePatRaw.Add(string.Join(":", force));
            }
            else if (text == "prepatdisconnectpins")
            {
                List<string> force = arrstr.ToList();
                force.RemoveAt(0);
                nHardIpInfo.DisconnectPrePat.Add(string.Join(",", force));
            }
            else if (text == "limits")
            {
                nHardIpInfo.SpecialLimits.Add(param);
            }
            else
            {
                ApplyMiscTextFieldExtra(nHardIpInfo, text, arrstr, tmpstr);
            }
        }

        private void ApplyMiscTextFieldExtra(HardIpInfo nHardIpInfo, string text, string[] arrstr, string tmpstr)
        {
            if (text == "newseqmeasstorename")
            {
                nHardIpInfo.NewInfo.MeasStoreName = arrstr[1];
            }
            else if (text == "newseqcalcequname")
            {
                nHardIpInfo.NewInfo.CalcEquation = arrstr[1];
            }
            else if (text == "newseqcalcstorename")
            {
                nHardIpInfo.NewInfo.CalcStoreName = arrstr[1];
            }
            else if (text == "newseqmeasrange")
            {
                nHardIpInfo.NewInfo.MeasRange = arrstr[1];
            }
            else if (text.Contains("newseqexpectfrequency"))
            {
                nHardIpInfo.NewInfo.ExpectFreq = arrstr[1];
            }
            else if (text == "trimloophlimit")
            {
                nHardIpInfo.NewInfo.TrimLoopHLimit.Add(arrstr[1]);
            }
            else if (text == "trimloopllimit")
            {
                nHardIpInfo.NewInfo.TrimLoopLLimit.Add(arrstr[1]);
            }
            else if (text.Contains("newseqinterposepremeas"))
            {
                nHardIpInfo.MiscInfo = new List<string> { "Interpose_PreMeas:" + arrstr[1] };
            }
            else if (text.Contains("newseqinterposepostmeas"))
            {
                nHardIpInfo.MiscInfo = new List<string> { "InterposePostMeas:" + arrstr[1] };
            }
            else if (text == "sendlsbfirst")
            {
                nHardIpInfo.IsSendLsbFirstInfo = arrstr[1];
            }
            else if (text == "caplsbfirst")
            {
                nHardIpInfo.IsCapLsbFirstInfo = arrstr[1];
            }
            else if (text.Contains("poweronfly"))
            {
                ApplyPowerOnFly(nHardIpInfo, arrstr);
            }
            else if (text.Contains("capstorename"))
            {
                nHardIpInfo.CapStoreName = tmpstr;
            }
            else if (text.Contains("psets"))
            {
                nHardIpInfo.PsetUseStr = tmpstr;
            }
        }

        private void ApplyPowerOnFly(HardIpInfo nHardIpInfo, string[] arrstr)
        {
            try
            {
                if (arrstr.Length != 2)
                {
                    string second = arrstr[1].ToLower();
                    if (second.Contains("vs_main"))
                    {
                        nHardIpInfo.PowerOnFlyItem.VMain = arrstr[2].Split(',').ToList();
                    }
                    else if (second.Contains("vs_alt"))
                    {
                        nHardIpInfo.PowerOnFlyItem.VAlt = arrstr[2].Split(',').ToList();
                    }
                }
                else
                {
                    nHardIpInfo.PowerOnFlyItem = new PowerOnFly { SwitchSeq = arrstr[1] };
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private List<HardIpSeqInfo> ExtractMeasSeqInfoNew(string measSeqStr, Dictionary<string, List<Dictionary<string, List<string>>>> measSet, string measNameStr)
        {
            var mMeasSeqInfo = new List<HardIpSeqInfo>();
            if (string.IsNullOrEmpty(measSeqStr))
            {
                return mMeasSeqInfo;
            }

            List<string> measNameList = measNameStr.Split('+').ToList();
            int seqCount = -1;
            var seqBox = new Dictionary<string, int>();
            foreach (string seq in measSeqStr.Split(','))
            {
                seqCount++;
                string lMeasName = measNameList.Count > 1 ? measNameList[seqCount] : measNameList[0];

                try
                {
                    if (!seqBox.ContainsKey(seq))
                    {
                        seqBox.Add(seq, seqCount);
                    }

                    seqBox[seq] = seqCount;

                    var nMeasSeqInfo = new HardIpSeqInfo();
                    mMeasSeqInfo.Add(nMeasSeqInfo);
                    nMeasSeqInfo.SeqName = seq;
                    nMeasSeqInfo.MeasName = lMeasName;

                    Dictionary<string, List<string>> specifyPins = seqBox[seq] < measSet[seq.ToLower()].Count
                        ? measSet[seq.ToLower()][seqBox[seq]]
                        : measSet[seq.ToLower()][0];

                    foreach (KeyValuePair<string, List<string>> pinGroup in specifyPins)
                    {
                        foreach (string pin in pinGroup.Value)
                        {
                            nMeasSeqInfo.PinList.Add(pin);
                            var measPin = new MeasPin(pin, "Meas" + seq)
                            {
                                SequenceIndex = seqCount + 1,
                                OriPinName = pinGroup.Key
                            };
                            nMeasSeqInfo.MeasPins.Add(measPin);
                        }
                    }
                }
                catch (Exception)
                {
                    //pass
                }
            }
            return mMeasSeqInfo;
        }

        public HardIpInfo ConvertHardIpInfoNewToOld(HardIpInfo oldInfo, HardIpInfoNew newInfo)
        {
            oldInfo.MeasSeqStr = newInfo.MeasSeq.Replace('|', ',');
            oldInfo.MeasName = newInfo.MeasName.Replace('|', '+').Replace('>', ',');
            oldInfo.CalcStoreName = newInfo.CalcStoreName;

            string[] measSeq = newInfo.MeasSeq.Split('|');
            List<string> measPins = newInfo.MeasPin.Split('|').ToList();
            oldInfo.MeasVStr = "";
            oldInfo.MeasIStr = "";
            oldInfo.MeasFStr = "";
            oldInfo.MeasVdiffStr = "";
            oldInfo.MeasVdiff2Str = "";
            oldInfo.MeasIdiffStr = "";
            oldInfo.MeasFdiffStr = "";
            oldInfo.MeasVocmStr = "";
            oldInfo.MeasR1Str = "";
            oldInfo.MeasR2Str = "";
            oldInfo.MeasDutyCycleStr = "";
            oldInfo.MeasVdmStr = "";
            var vStr = new List<string>();
            var iStr = new List<string>();
            var fStr = new List<string>();
            var vdiffStr = new List<string>();
            var vdiff2Str = new List<string>();
            var idiffStr = new List<string>();
            var fdiffStr = new List<string>();
            var vocmStr = new List<string>();
            var r1Str = new List<string>();
            var r2Str = new List<string>();
            var dutyCycleStr = new List<string>();
            var vdmStr = new List<string>();

            for (int i = 0; i < measSeq.Length; i++)
            {

                string seq = measSeq[i].ToLower();
                if (!oldInfo.MeaSet.ContainsKey(seq))
                {
                    oldInfo.MeaSet.Add(seq, new List<Dictionary<string, List<string>>>());
                }

                string names = (i >= measPins.Count ? measPins[0] : measPins[i]).Trim('>');
                Dictionary<string, List<string>> item = names.Split(',').ToDictionary(name => name, name => new List<string> { name });
                oldInfo.MeaSet[seq].Add(item);
                AccumulateMeasSeq(oldInfo, seq, names, measPins[i], vStr, iStr, fStr, vdiffStr, vdiff2Str, idiffStr, fdiffStr, vocmStr, r1Str, r2Str, dutyCycleStr, vdmStr);
            }

            oldInfo.MeasVStr = oldInfo.MeasVStr.TrimEnd(',');
            oldInfo.MeasIStr = oldInfo.MeasIStr.TrimEnd(',');
            oldInfo.MeasFStr = oldInfo.MeasFStr.TrimEnd('+');
            oldInfo.MeasVdiffStr = oldInfo.MeasVdiffStr.TrimEnd(',');
            oldInfo.MeasVdiff2Str = oldInfo.MeasVdiff2Str.TrimEnd(',');
            oldInfo.MeasIdiffStr = oldInfo.MeasIdiffStr.TrimEnd(',');
            oldInfo.MeasFdiffStr = oldInfo.MeasFdiffStr.TrimEnd(',');
            oldInfo.MeasVocmStr = oldInfo.MeasVocmStr.TrimEnd(',');
            oldInfo.MeasR1Str = oldInfo.MeasR1Str.TrimEnd(',');
            oldInfo.MeasR2Str = oldInfo.MeasR2Str.TrimEnd(',');
            oldInfo.MeasDutyCycleStr = oldInfo.MeasDutyCycleStr.TrimEnd(',');
            oldInfo.MeasVdmStr = oldInfo.MeasVdmStr.TrimEnd(',');

            return oldInfo;
        }

        private static void AccumulateMeasSeq(HardIpInfo oldInfo, string seq, string names, string measPin,
            List<string> vStr, List<string> iStr, List<string> fStr, List<string> vdiffStr, List<string> vdiff2Str,
            List<string> idiffStr, List<string> fdiffStr, List<string> vocmStr, List<string> r1Str, List<string> r2Str,
            List<string> dutyCycleStr, List<string> vdmStr)
        {
            if (seq == "v")
            {
                if (!vStr.Exists(x => x == names))
                {
                    vStr.Add(measPin);
                    oldInfo.MeasVStr += names + ",";
                }
            }
            else if (seq == "i")
            {
                if (!iStr.Exists(x => x == names))
                {
                    iStr.Add(measPin);
                    oldInfo.MeasIStr += names + ",";
                }
            }
            else if (seq == "f")
            {
                if (!fStr.Exists(x => x == names))
                {
                    fStr.Add(names);
                    oldInfo.MeasFStr += names + "+";
                }
            }
            else if (seq == "vdiff")
            {
                if (!vdiffStr.Exists(x => x == names))
                {
                    vdiffStr.Add(names);
                    oldInfo.MeasVdiffStr += names + ",";
                }
            }
            else if (seq == "vdiff2")
            {
                if (!vdiff2Str.Exists(x => x == names))
                {
                    vdiff2Str.Add(measPin);
                    oldInfo.MeasVdiff2Str += names + ",";
                }
            }
            else if (seq == "idiff")
            {
                if (!idiffStr.Exists(x => x == names))
                {
                    idiffStr.Add(measPin);
                    oldInfo.MeasIdiffStr += names + ",";
                }
            }
            else
            {
                AccumulateMeasSeqExtra(oldInfo, seq, names, measPin, fdiffStr, vocmStr, r1Str, r2Str, dutyCycleStr, vdmStr);
            }
        }

        private static void AccumulateMeasSeqExtra(HardIpInfo oldInfo, string seq, string names, string measPin,
            List<string> fdiffStr, List<string> vocmStr, List<string> r1Str, List<string> r2Str,
            List<string> dutyCycleStr, List<string> vdmStr)
        {
            if (seq == "fdiff")
            {
                if (!fdiffStr.Exists(x => x == names))
                {
                    fdiffStr.Add(measPin);
                    oldInfo.MeasFdiffStr += names + ",";
                }
            }
            else if (seq == "vocm")
            {
                if (!vocmStr.Exists(x => x == names))
                {
                    vocmStr.Add(measPin);
                    oldInfo.MeasVocmStr += names + ",";
                }
            }
            else if (seq == "r1")
            {
                if (!r1Str.Exists(x => x == names))
                {
                    r1Str.Add(measPin);
                    oldInfo.MeasR1Str += names + ",";
                }
            }
            else if (seq == "r2")
            {
                if (!r2Str.Exists(x => x == names))
                {
                    r2Str.Add(measPin);
                    oldInfo.MeasR2Str += names + ",";
                }
            }
            else if (seq == "dutycycle")
            {
                if (!dutyCycleStr.Exists(x => x == names))
                {
                    dutyCycleStr.Add(measPin);
                    oldInfo.MeasDutyCycleStr += names + ",";
                }
            }
            else if (seq == "vdm")
            {
                if (!vdmStr.Exists(x => x == names))
                {
                    vdmStr.Add(names);
                    oldInfo.MeasVdmStr += names + ",";
                }
            }
        }

        public List<string> DecompGroups(string pinGroup)
        {
            var newPinList = new List<string>();
            if (PinGroupList.ContainsKey(pinGroup))
            {
                foreach (string pin in PinGroupList[pinGroup])
                {
                    if (pin.Equals(pinGroup, StringComparison.CurrentCultureIgnoreCase))
                    {
                        newPinList.Add(pin);
                        continue;
                    }
                    newPinList.AddRange(DecompGroups(pin));
                }
            }
            else
            {
                newPinList.Add(pinGroup);
            }
            return newPinList;
        }

        public string ConvertToNetName(string measStr)
        {
            string resultStr;
            if (!measStr.Contains("VDD", StringComparison.OrdinalIgnoreCase))
            {
                resultStr = measStr;
            }
            else
            {
                string newMeasStr = "";
                foreach (string seq in measStr.Split('+'))
                {
                    string newSeq = "";
                    if (seq != "")
                    {
                        var newSeqList = new List<string>();
                        foreach (string pinName in seq.Split(','))
                        {
                            string cpNetName = "";
                            string ftNetName = "";

                            //KIMI test
                            PowerMerge.GetCpFtNetName(pinName, ref cpNetName, ref ftNetName);
                            if (cpNetName == "" && ftNetName == "")
                            {
                                newSeqList.Add(pinName);
                            }
                            else if (cpNetName == ftNetName)
                            {
                                newSeqList.Add(cpNetName);

                            }
                            else if (cpNetName != ftNetName)
                            {
                                if (cpNetName != "" && cpNetName != "N/A" && !newSeqList.Contains(cpNetName))
                                {
                                    newSeqList.Add("CP=" + cpNetName.Replace(",", ",CP="));
                                }
                                if (ftNetName != "" && ftNetName != "N/A" && !newSeqList.Contains(ftNetName))
                                {
                                    newSeqList.Add("FT=" + ftNetName.Replace(",", ",FT="));
                                }
                            }
                        }
                        newSeq = string.Join(",", newSeqList.Distinct());
                    }
                    newMeasStr += newSeq + "+";
                }
                resultStr = newMeasStr.Remove(newMeasStr.Length - 1, 1);
            }
            return resultStr;
        }
    }
}
