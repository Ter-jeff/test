using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class MeasPinsAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        int seqCount = ctx.HardIpInfo.NewInfo != null ? ctx.HardIpInfo.NewInfo.SeqInfo.Count :
                ctx.HardIpInfo.SeqInfo.Count == 0 ? ctx.Pattern.TestPlanSequences.Count : ctx.HardIpInfo.SeqInfo.Count;
        if (seqCount <= 0)
        {
            return;
        }
        #region Gets measPins from TestSequence

        var measVPins = new List<string>();
        var measIPins = new List<string>();
        var measFPins = new List<string>();
        var measFdiffPins = new List<string>();

        for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
        {
            CollectMeasPinsBySequence(ctx, sequenceIndex, measVPins, measIPins, measFPins, measFdiffPins);
        }
        #endregion

        //MeasureV_PinS
        if (measVPins.Any(p => !string.IsNullOrEmpty(p)))
        {
            string measV = DataConvertor.RemoveDummyPlusSign(DataConvertor.SortCpFtPin(
                SearchInfo.CheckInfoByStoreName(string.Join("+", measVPins), ctx.OriginalStoreName, '+')));
            if (measV.Length >= 6000)
            {
                measV = WriteMeaspinToRegAssign(ctx.Pattern, ctx.TestSequence, measVPins, "MeasV", ctx.InputData);
            }
            function.SetParamValue("measureVPins", measV);
        }
        //MeasureF_PinS_SingleEnd
        if (measFPins.Any(p => !string.IsNullOrEmpty(p)))
        {
            string measF = DataConvertor.RemoveDummyPlusSign(DataConvertor.SortCpFtPin(
                SearchInfo.CheckInfoByStoreName(string.Join("+", measFPins), ctx.OriginalStoreName, '+')));
            if (measF.Length >= 6000)
            {
                measF = WriteMeaspinToRegAssign(ctx.Pattern, ctx.TestSequence, measFPins, "MeasF", ctx.InputData);
            }
            function.SetParamValue("measureFPins", measF);
        }
        //MeasureF_PinS_Differtial
        if (measFdiffPins.Any(p => !string.IsNullOrEmpty(p)))
        {
            function.SetParamValue("measureFDifferentialPins", DataConvertor.RemoveDummyPlusSign(DataConvertor.SortCpFtPin(
                SearchInfo.CheckInfoByStoreName(string.Join("+", measFdiffPins), ctx.OriginalStoreName, '+'))));
        }

        //MeasureI_pinS
        if (measIPins.Any(p => !string.IsNullOrEmpty(p)))
        {
            //function.SetParamValue("MeasI_pinS", DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(DataConvertor.RemoveDummyPlusSign(string.Join("+", measIPins)), storeNameOri, '+')));
            string measI = DataConvertor.RemoveDummyPlusSign(DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(string.Join("+", measIPins), ctx.OriginalStoreName, '+')));
            if (measI.Length >= 6000)
            {
                measI = WriteMeaspinToRegAssign(ctx.Pattern, ctx.TestSequence, measIPins, "MeasI", ctx.InputData);
            }
            function.SetParamValue("measureIPins", measI);

            //MeasI_Range
            //currentRange = SearchInfo.GetIRangeByJob(currentRange);
            //function.SetParamValue("MeasI_Range", DataConvertor.RedefineRange(pattern, info, SearchInfo.CheckInfoByStoreName(SearchInfo.GetIRange(pattern, info, voltage), storeNameOri, '|')));
            function.SetParamValue("measureRange", SearchInfo.CheckInfoByStoreName(SearchInfo.GetIRange(ctx.Pattern, ctx.HardIpInfo, ctx.Voltage), ctx.OriginalStoreName, '|'));
        }
    }

    private void CollectMeasPinsBySequence(FunctionAssignContext ctx, int sequenceIndex, List<string> measVPins, List<string> measIPins,
        List<string> measFPins, List<string> measFdiffPins)
    {
        var seqPins = ctx.Pattern.MeasPins.Where(p => p.SequenceIndex == sequenceIndex).ToList();
        bool isMeasR = seqPins.Exists(x => x.MeasType == MeasType.MeasR1 || x.MeasType == MeasType.MeasR2);
        ForceConditionType mode = isMeasR ? SearchInfo.GetForceMode(seqPins) : ForceConditionType.Normal;
        string selVType =
            seqPins.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase)) ? MeasType.MeasVdiff :
            MeasType.MeasV;
        bool isPinGroup = seqPins.Exists(p => TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(p.PinName));
        var measIInter = new List<MeasPin>();
        var measVInter = new List<MeasPin>();

        if (mode == ForceConditionType.FiMode) //MeasR Force I mode
        {
            measIInter =
            seqPins.Where(
                p => p.MeasType.Equals(MeasType.MeasI, StringComparison.OrdinalIgnoreCase)).ToList();
            measVInter = seqPins.Where(p =>
            p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase) ||
            p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase) || p.MeasType.Equals(MeasType.MeasR1, StringComparison.OrdinalIgnoreCase) ||
                    p.MeasType.Equals(MeasType.MeasR2, StringComparison.OrdinalIgnoreCase)
            ).ToList();
        }
        else if (mode == ForceConditionType.FvMode) // MeasR Force V mode
        {
            measIInter =
            seqPins.Where(
                p => p.MeasType.Equals(MeasType.MeasI, StringComparison.OrdinalIgnoreCase)
                     || p.MeasType.Equals(MeasType.MeasR1, StringComparison.OrdinalIgnoreCase) ||
                     p.MeasType.Equals(MeasType.MeasR2, StringComparison.OrdinalIgnoreCase)).ToList();
            measVInter = seqPins.Where(p =>
            p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase) ||
            p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase)
            ).ToList();
        }
        else
        {
            measIInter =
            seqPins.Where(p =>
            p.MeasType.Equals(MeasType.MeasI, StringComparison.OrdinalIgnoreCase) ||
            p.MeasType.Equals(MeasType.MeasVdiff2, StringComparison.OrdinalIgnoreCase)
            ).ToList();

            measVInter = seqPins.Where(
                p => p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase)
                ).ToList();
        }
        var measFInter = seqPins.Where(
            p => p.MeasType.Equals(MeasType.MeasF, StringComparison.OrdinalIgnoreCase) || p.MeasType.Equals(MeasType.MeasD)).ToList();
        var measFdiffInter = seqPins.Where(
            p => p.MeasType.Equals(MeasType.MeasFdiff, StringComparison.OrdinalIgnoreCase)).ToList();

        measVPins.Add(isPinGroup ? string.Join(",", RearrangePin(measVInter)) :
            string.Join(",", measVInter.Select(p =>
            p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase) ?
            p.PinName : SearchInfo.GenDiffGroupName(p.PinName, false))));

        if (measIInter.Any(p => p.MeasType == MeasType.MeasVdiff2) && measIInter.SelectMany(x => x.ForceConditions).Any()) //Vdiff2 meas pins from force for single pin limit by Orange/HungYu
        {
            measIPins.Add(string.Join(",", measIInter.SelectMany(p => p.ForceConditions.SelectMany(q => q.ForcePins.Select(s => s.PinName))).Distinct()));
        }
        else
        {
            measIPins.Add(string.Join(",", measIInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, false))));
        }

        measFPins.Add(string.Join(",", measFInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, true))));
        measFdiffPins.Add(string.Join(",", measFdiffInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, true))));
    }

    private static List<string> RearrangePin(List<MeasPin> pins)
    {
        var forcePins = pins.SelectMany(p => p.ForceConditions).SelectMany(p => p.ForcePins).Select(p => p.PinName).Distinct().ToList();
        var measPins = pins.Select(p => p.PinName).Distinct().ToList();
        if (forcePins.Count == 0)
        {
            return measPins;
        }

        if (forcePins.Except(measPins).Count() == 0 &&
            measPins.Except(forcePins).Count() == 0)
        {
            return measPins;
        }

        var totalMeasPins = new List<string>();
        foreach (string meas in measPins)
        {
            totalMeasPins.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(meas));
        }
        var totalForcePins = new List<string>();
        foreach (string force in forcePins)
        {
            totalForcePins.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(force));
        }
        if (totalForcePins.Except(totalMeasPins).Count() == 0 &&
            totalMeasPins.Except(totalForcePins).Count() == 0)
        {
            return forcePins;
        }
        return pins.Select(p => p.PinName).ToList();
    }

    private static string WriteMeaspinToRegAssign(HardIpPattern pattern, string testSequence, List<string> pinList, string type, HardIpInputData inputData)
    {
        var hardIpRegAssign = new HardIpRegAssign();
        string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
        hardIpRegAssign.SubBlockName = blockName + "_" + CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
        if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
        {
            hardIpRegAssign.SubBlockName += "_DD";
        }

        switch (type)
        {
            case "MeasV":
                hardIpRegAssign.Type = RegisterAssignType.MeasV_PinS;
                break;
            case "MeasI":
                hardIpRegAssign.Type = RegisterAssignType.MeasI_PinS;
                break;
            case "MeasF":
                hardIpRegAssign.Type = RegisterAssignType.MeasF_PinS_SingleEnd;
                break;
        }

        if (!inputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(hardIpRegAssign.SubBlockName, StringComparison.CurrentCultureIgnoreCase) &&
                                                          x.Type == hardIpRegAssign.Type))
        {
            List<List<string>> regAssginList = new List<List<string>>();
            List<string> testSequenceArr = testSequence.Split(',').ToList();
            if (testSequenceArr.Count == pinList.Count)
            {
                for (int index = 0; index < pinList.Count; index++)
                {
                    string measPin = pinList[index];
                    measPin = measPin.Replace(",FT=", ";FT=");
                    List<string> data = new List<string> { testSequenceArr.ElementAt(index) };
                    data.AddRange(new List<string> { measPin });
                    regAssginList.Add(data);
                }
            }

            hardIpRegAssign.RegAssignList = regAssginList;
            inputData.HardIpRegAssigns.Add(hardIpRegAssign);

        }
        return $"Reg_assign:{hardIpRegAssign.SubBlockName}";
    }
}
