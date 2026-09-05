using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal class WiTrimBestCodeGenerator(TemplateAutoGen templateAutoGen)
    {
        private readonly TemplateAutoGen _owner = templateAutoGen;

        public void GenPlanWithWiTrimItem(List<TemplateRow> templateRows, HardIpInfo hardIpInfo, bool isReadCap, int blockindex, ref int stepIndex)
        {
            GenerateWiTrimRfItemRows(hardIpInfo, templateRows, blockindex, ref stepIndex);
            GenerateWiTrimMeasCRowIfNeeded(hardIpInfo, templateRows, blockindex, stepIndex, isReadCap);

            PrepareWiTrimSegments(hardIpInfo, out List<string> trimSegs, out List<int> trimBits, out List<string> trimFuses);

            GenerateBestCodeRows(templateRows, trimSegs, trimBits, trimFuses, blockindex, stepIndex);
            GenerateBestValueRows(templateRows, hardIpInfo, trimSegs, trimFuses, blockindex, stepIndex, isReadCap);
            GenerateBestValueDiffRows(templateRows, hardIpInfo, trimSegs, trimFuses, blockindex, stepIndex, isReadCap);
            GenerateVerificationCodeRows(templateRows, trimSegs, trimBits, trimFuses, blockindex, stepIndex);
            GenerateVerificationValueRows(templateRows, hardIpInfo, trimSegs, trimFuses, blockindex, stepIndex, isReadCap);
            GenerateVerificationValueDiffRows(templateRows, hardIpInfo, trimSegs, trimFuses, blockindex, stepIndex, isReadCap);
        }

        private void GenerateWiTrimRfItemRows(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, int blockindex, ref int stepIndex)
        {
            if (hardIpInfo.NewInfo == null)
            {
                return;
            }

            bool isRFItem = hardIpInfo.NewInfo.SeqInfo.SelectMany(p => p.MeasPins).ToList()
                .Exists(p => TemplateAutoGen.MyRegex3().IsMatch(p.MeasType)) ||
                hardIpInfo.MiscInfo.Exists(p => TemplateAutoGen.MyRegex4().IsMatch(p));
            if (!isRFItem)
            {
                return;
            }

            int i = 0;
            foreach (HardIpSeqInfoNew seq in hardIpInfo.NewInfo.SeqInfo)
            {
                switch (seq.MeasSeq.ToUpper())
                {
                    default:
                        _owner.GenPlanWithNonVdiffNew(hardIpInfo.Payload, templateRows, seq, i.ToString(), blockindex, stepIndex, isRFItem);
                        TemplateAutoGenHelpers1.GenerateCalcEquation(templateRows, seq, i.ToString(), blockindex, stepIndex);
                        break;
                }

                i++;
                stepIndex++;
            }
        }

        private static void GenerateWiTrimMeasCRowIfNeeded(HardIpInfo hardIpInfo, List<TemplateRow> templateRows, int blockindex, int stepIndex, bool isReadCap)
        {
            string regMeasC = @"\(\w+\)";
            if (!isReadCap && hardIpInfo.DsscOut.Length != 0 && !Regex.IsMatch(hardIpInfo.TrimTarget, regMeasC, RegexOptions.IgnoreCase) && LocalSpecs.Options.Device != EnumDevice.LCD)
            {
                MeasCTemplateItemGenerator.GenMeasCTemplateItemNew(hardIpInfo, templateRows, blockindex, stepIndex);
            }
        }

        private static void PrepareWiTrimSegments(HardIpInfo hardIpInfo, out List<string> trimSegs, out List<int> trimBits, out List<string> trimFuses)
        {
            string? trimMeasName = hardIpInfo.MiscInfo.SelectMany(p => p.Split(';')).FirstOrDefault(p => p.Split(':')[0].EqualsIgnoreCase("TrimMeasName"));
            if (!string.IsNullOrEmpty(trimMeasName))
            {
                trimMeasName = trimMeasName.Split(':')[1].Trim(';');
            }
            trimSegs = [.. (trimMeasName ?? "").Split('_')];
            //trimSegs[4] = "BSTC";//Test Name fix for draco 260302
            trimSegs[3] = "BSTC";
            trimBits = !string.IsNullOrEmpty(hardIpInfo.TrimFuseName) ? TemplateAutoGenHelpers1.GetTrimRegBit(hardIpInfo) : [];
            if (trimBits.Count == 0)
            {
                string regTrim = @"Trimbits:(?<bits>[\d\,]+)";
                string bits = Regex.Match(string.Join(";", hardIpInfo.MiscInfo), regTrim, RegexOptions.IgnoreCase).Groups["bits"].Value;
                if (!string.IsNullOrEmpty(bits))
                {
                    trimBits = [.. bits.Split(',').Select(int.Parse)];
                }
            }
            trimSegs[5] = trimBits.Count == 2 ||
                (trimBits.Count != 2 && hardIpInfo.MeasName.Split('+').Length == 2) ||
                hardIpInfo.BestCodeCalcFunc.Contains("algAvgSeq1PlusSeq2") ?
                "X" : trimSegs[5];
            trimFuses = [.. hardIpInfo.TrimFuseName.Split(',').Select(tfn => tfn.TrimEnd(';').Replace("_", "").ToUpper())];
        }

        private static void GenerateBestCodeRows(List<TemplateRow> templateRows, List<string> trimSegs, List<int> trimBits, List<string> trimFuses, int blockindex, int stepIndex)
        {
            WirelessTemplateRow newTempRow;
            int index = -1;
            if (trimBits.Count > 0)
            {
                foreach (int bit in trimBits)
                {
                    index++;

                    string fusename = trimBits.Count >= 2 ? trimFuses[index] : trimSegs[3];
                    //trimSegs[3] = fusename.Replace("_", "").ToUpper(); //Test Name Ori 260304
                    //Test Name fix for draco 260302
                    trimSegs[4] = "X";
                    //Test Name fix for draco 260302
                    trimSegs[5] = "X";
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BestCode", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    newTempRow.LoLimit.Add("CP1", "0");
                    newTempRow.HiLimit.Add("CP1", (Math.Pow(2, bit) - 1).ToString());
                }
            }
        }

        private static void GenerateBestValueRows(List<TemplateRow> templateRows, HardIpInfo hardIpInfo, List<string> trimSegs, List<string> trimFuses, int blockindex, int stepIndex, bool isReadCap)
        {
            //trimSegs[4] = "BSTV";//Test Name fix for draco
            trimSegs[3] = "BSTV";
            WirelessTemplateRow newTempRow;
            if (hardIpInfo.BestCodeCalcFunc.EqualsIgnoreCase("algParallel2TrimFindBestCode"))
            {
                foreach (string trimfuse in trimFuses)
                {
                    //trimSegs[3] = trimfuse.Replace("_", "").ToUpper();//Test Name Ori 260304
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BestValue", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
                }
            }
            else
            {
                //trimSegs[3] = TrimBits.Count() >= 2 ? string.Join("-", trimFuses) : trimSegs[3]; //Test Name Ori 260304
                newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BestValue", string.Join("_", trimSegs));
                templateRows.Add(newTempRow);
                TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
            }
        }

        private static void GenerateBestValueDiffRows(List<TemplateRow> templateRows, HardIpInfo hardIpInfo, List<string> trimSegs, List<string> trimFuses, int blockindex, int stepIndex, bool isReadCap)
        {
            if (hardIpInfo.BestCodeCalcFunc.Contains("algFindBestCodeWindow"))
            {
                return;
            }

            WirelessTemplateRow newTempRow;
            if (hardIpInfo.BestCodeCalcFunc.EqualsIgnoreCase("algParallel2TrimFindBestCode"))
            {
                //trimSegs[4] = "BSTV-TrimTarget-DIFF";//Test Name fix for draco
                trimSegs[3] = "BSTV-TrimTarget-DIFF";
                foreach (string trimfuse in trimFuses)
                {
                    //trimSegs[3] = trimfuse.Replace("_", "").ToUpper();//Test Name Ori 260304
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BSTVTrimTargetDIFF", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
                }

                //trimSegs[4] = "BSTV-TrimTarget-DIFF-PERCENT";//Test Name fix for draco
                trimSegs[3] = "BSTV-TrimTarget-DIFF-PERCENT";
                foreach (string trimfuse in trimFuses)
                {
                    trimSegs[3] = trimfuse.Replace("_", "").ToUpper();
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BSTVTrimTargetDIFFPERCENT", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
                }
            }
            else
            {
                //trimSegs[3] = TrimBits.Count() >= 2 ? string.Join("-", trimFuses) : trimSegs[3];//Test Name Ori 260304

                //trimSegs[4] = "BSTV-TrimTarget-DIFF";//Test Name fix for draco
                trimSegs[3] = "BSTV-TrimTarget-DIFF";
                newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BSTVTrimTargetDIFF", string.Join("_", trimSegs));
                templateRows.Add(newTempRow);
                TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);

                //trimSegs[4] = "BSTV-TrimTarget-DIFF-PERCENT"; //Test Name fix for draco
                trimSegs[3] = "BSTV-TrimTarget-DIFF-PERCENT";
                newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for BSTVTrimTargetDIFFPERCENT", string.Join("_", trimSegs));
                templateRows.Add(newTempRow);
                TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
            }
        }

        private static void GenerateVerificationCodeRows(List<TemplateRow> templateRows, List<string> trimSegs, List<int> trimBits, List<string> trimFuses, int blockindex, int stepIndex)
        {
            //trimSegs[4] = "VRFC"; //Test Name fix for draco
            trimSegs[3] = "VRFC";
            if (trimBits.Count > 0)
            {
                int index = -1;
                foreach (int bit in trimBits)
                {
                    index++;

                    string fusename = trimBits.Count >= 2 ? trimFuses[index] : trimSegs[3];
                    //trimSegs[3] = fusename.Replace("_", "").ToUpper(); //Test Name Ori 260304
                    var newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VerificationCode", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    newTempRow.LoLimit.Add("CP1", "0");
                    newTempRow.HiLimit.Add("CP1", (Math.Pow(2, bit) - 1).ToString());
                }
            }
        }

        private static void GenerateVerificationValueRows(List<TemplateRow> templateRows, HardIpInfo hardIpInfo, List<string> trimSegs, List<string> trimFuses, int blockindex, int stepIndex, bool isReadCap)
        {
            //trimSegs[4] = "VRFV"; //Test Name fix for draco
            trimSegs[3] = "VRFV";
            WirelessTemplateRow newTempRow;
            if (hardIpInfo.BestCodeCalcFunc.EqualsIgnoreCase("algParallel2TrimFindBestCode"))
            {
                foreach (string trimfuse in trimFuses)
                {
                    //trimSegs[3] = trimfuse.Replace("_", "").ToUpper();//Test Name Ori 260304
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VerificationValue", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
                }
            }
            else
            {
                //trimSegs[3] = TrimBits.Count() >= 2 ? string.Join("-", trimFuses) : trimSegs[3];//Test Name Ori 260304
                newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VerificationValue", string.Join("_", trimSegs));
                templateRows.Add(newTempRow);
                TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
            }
        }

        private static void GenerateVerificationValueDiffRows(List<TemplateRow> templateRows, HardIpInfo hardIpInfo, List<string> trimSegs, List<string> trimFuses, int blockindex, int stepIndex, bool isReadCap)
        {
            if (hardIpInfo.BestCodeCalcFunc.Contains("algFindBestCodeWindow"))
            {
                return;
            }

            WirelessTemplateRow newTempRow;
            if (hardIpInfo.BestCodeCalcFunc.EqualsIgnoreCase("algParallel2TrimFindBestCode"))
            {
                //trimSegs[4] = "VRFV-VRFTarget-DIFF"; //Test Name fix for draco
                trimSegs[3] = "VRFV-VRFTarget-DIFF";
                foreach (string trimfuse in trimFuses)
                {
                    trimSegs[3] = trimfuse.Replace("_", "").ToUpper();
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VRFVVRFTargetDiff", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
                }

                //trimSegs[4] = "VRFV-VRFTarget-DIFF-PERCENT"; //Test Name fix for draco
                trimSegs[3] = "VRFV-VRFTarget-DIFF-PERCENT";
                foreach (string trimfuse in trimFuses)
                {
                    trimSegs[3] = trimfuse.Replace("_", "").ToUpper();
                    newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VRFVVRFTargetDiffPercent", string.Join("_", trimSegs));
                    templateRows.Add(newTempRow);
                    TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
                }
            }
            else
            {
                //trimSegs[3] = TrimBits.Count() >= 2 ? string.Join("-", trimFuses) : trimSegs[3];//Test Name Ori 260304

                //trimSegs[4] = "VRFV-VRFTarget-DIFF"; //Test Name fix for draco
                trimSegs[3] = "VRFV-VRFTarget-DIFF";
                newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VRFVVRFTargetDiff", string.Join("_", trimSegs));
                templateRows.Add(newTempRow);
                TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);

                //trimSegs[4] = "VRFV-VRFTarget-DIFF-PERCENT";//Test Name fix for draco
                trimSegs[3] = "VRFV-VRFTarget-DIFF-PERCENT";
                newTempRow = new WirelessTemplateRow(blockindex, blockindex + "." + stepIndex, "TestName for VRFVVRFTargetDiffPercent", string.Join("_", trimSegs));
                templateRows.Add(newTempRow);
                TemplateAutoGenHelpers1.GetBestValue(newTempRow, hardIpInfo.NewInfo, isReadCap);
            }
        }
    }
}
