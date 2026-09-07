using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner;

public record FunctionAssignContext
{
    public HardIpPattern Pattern { get; }
    public HardIpInfo HardIpInfo { get; }
    public string Voltage { get; }
    public HardIpInputData InputData { get; }
    public string OriginalStoreName { get; }
    public string MeasStoreName { get; }
    public string TestSequence { get; }
    public bool IsAddSweep { get; }
    public string DigSrcAssignmentVal { get; }

    public FunctionAssignContext(HardIpPattern pattern, string voltage, HardIpInputData inputData)
    {
        Pattern = pattern;
        HardIpInfo = HardIpService.GetHardIpInfo(pattern);
        Voltage = voltage;
        InputData = inputData;

        bool isAddSweep = false;
        DigSrcAssignmentVal = string.Join(";", Pattern.ProcessSweepData(ref isAddSweep)
            .Split(';', '|')
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(x => !string.IsNullOrEmpty(x))
        );
        IsAddSweep = isAddSweep;

        string measSeq = GetMeasSeq(SearchInfo.GetMeasSequence(pattern).ToUpper());
        string storeNameOri = "";
        MeasStoreName = DataConvertor.SortCpFtPin(SearchInfo.GetStoreName(pattern, InputData, ref storeNameOri, measSeq));
        OriginalStoreName = storeNameOri;
        TestSequence = SearchInfo.CheckInfoByStoreName(measSeq, storeNameOri, ',', true);
    }

    private static string GetMeasSeq(string measSeq)
    {
        //FDIFF need to be changed to F in meas sequence

        var allseqs = new List<string>();
        foreach (string seq in measSeq.Split(','))
        {
            if (seq.Equals("VDIFF2"))
            {
                allseqs.Add(seq);
            }
            else
            {
                string replacedSeq = seq
                    .Replace("R1", "R")
                    .Replace("IDIFF", "I")
                    .Replace("R2", "Z")
                    .Replace("VDIFF", "V");
                allseqs.Add(replacedSeq);
            }
        }

        return string.Join(",", allseqs);
    }
}
