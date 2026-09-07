using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Const;

namespace CommonLib.Utility.FrcCalc
{
    public class FrcCalcMain
    {
        public static List<double> CalculateFrcFreq(List<int> dataList)
        {
            var sbcFreqCalculatorFrcList = new Dictionary<string, SbcSolutionFrc>();
            var totalSetOfFreq = new List<List<double>>();
            foreach (int data in dataList)
            {
                //represent to items of nWire
                var calculatorFrc = new SbcFreqCalculator();
                calculatorFrc.TargetFreq.Add(data);
                calculatorFrc.SolveSbcFreqFrc(out sbcFreqCalculatorFrcList);
                var currFreqSet = sbcFreqCalculatorFrcList.Select(p => p.Value.EngineList[0].PllInputFreq).ToList();
                totalSetOfFreq.Add(currFreqSet);
            }

            List<double> targetSets = totalSetOfFreq[0];
            foreach (List<double> list in totalSetOfFreq)
            {
                targetSets = [.. targetSets.Intersect(list)];
            }

            targetSets = [.. targetSets.OrderBy(x => x)];

            var sbCs = new List<double>();
            foreach (double pllFreq in targetSets)
            {
                KeyValuePair<string, SbcSolutionFrc> target = sbcFreqCalculatorFrcList.FirstOrDefault(p => Math.Abs(p.Value.EngineList[0].PllInputFreq - pllFreq) < CommonConst.Tolerance);
                sbCs.Add(target.Value.SbcFreq);
            }
            return sbCs;
        }
    }
}
