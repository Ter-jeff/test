using System.Collections.Generic;
using System.Linq;

namespace CommonLib.Utility.FrcCalc
{
    public class SbcFreqCalculator
    {
        private readonly double _patGenFreqHighLimit = 550000000;
        private readonly double _clkD8FreqLowLimit = 125000000;
        private readonly double _clkD8FreqHighLimit = 275000000;
        private readonly double _pdfFreqLowLimit = 3000000;
        private readonly double _pdfFreqHighLimit = 6000000;
        private readonly double _pllInputFreqHighLimit = 200000000;
        private readonly double _pllInputFreqLowLimit = 3000000;

        public string? Sdf { set; get; }
        public List<double> TargetFreq = [];

        private bool TrySingleValue(double targetValue, double pllInputFreq, out PaEngineItem paEngineItem)
        {
            paEngineItem = new PaEngineItem();
            double patGenFreq = 0.0;
            for (int hSpeedMode = 1; hSpeedMode <= 2; hSpeedMode++)
            {
                if (hSpeedMode == 1)
                {
                    patGenFreq = targetValue;
                }
                else if (hSpeedMode == 2)
                {
                    patGenFreq = targetValue / 2;
                }

                if (patGenFreq >= _patGenFreqHighLimit)
                {
                    continue;
                }

                int m2 = hSpeedMode;
                for (int x2 = 1; x2 <= 2; x2++)
                {
                    double clkD8Freq;
                    double pdfFreq;
                    int pdfHighRange;
                    int pdfLowRange;
                    if (x2 == 1)
                    {
                        int clkD8LowRange = (int)(_clkD8FreqLowLimit / patGenFreq);
                        int clkD8HighRange = (int)(_clkD8FreqHighLimit / patGenFreq);
                        // D2 = clkD8LowRange ~ clkD8HighRange
                        for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                        {
                            clkD8Freq = patGenFreq * i;
                            if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit || i > 65535)
                            {
                                continue;
                            }
                            pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                            pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                            for (int j = pdfLowRange; j <= pdfHighRange; j++)
                            {
                                pdfFreq = clkD8Freq / j;
                                if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                {
                                    continue;
                                }
                                if (pllInputFreq % pdfFreq == 0)
                                {
                                    paEngineItem.ClkD8Freq = clkD8Freq;
                                    paEngineItem.M2 = m2;
                                    paEngineItem.PatGenFreq = patGenFreq;
                                    paEngineItem.D2 = i;
                                    paEngineItem.X2 = x2;
                                    paEngineItem.PdfFreq = pdfFreq;
                                    paEngineItem.PllInputFreq = pllInputFreq;
                                    paEngineItem.M = j;
                                    paEngineItem.D1 = (int)(pllInputFreq / pdfFreq);
                                    paEngineItem.TargetFreq = targetValue;
                                    return true;
                                }
                            }
                        }

                    }
                    else if (x2 == 2) // D2 only can put 1
                    {
                        clkD8Freq = patGenFreq / 2;
                        // D2 = 1;
                        if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                        {
                            continue;
                        }
                        pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                        pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                        for (int j = pdfLowRange; j <= pdfHighRange; j++)
                        {
                            pdfFreq = clkD8Freq / j;
                            if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                            {
                                continue;
                            }
                            if (pllInputFreq % pdfFreq == 0)
                            {
                                paEngineItem.ClkD8Freq = clkD8Freq;
                                paEngineItem.M2 = m2;
                                paEngineItem.PatGenFreq = patGenFreq;
                                paEngineItem.D2 = 1;
                                paEngineItem.X2 = x2;
                                paEngineItem.PdfFreq = pdfFreq;
                                paEngineItem.PllInputFreq = pllInputFreq;
                                paEngineItem.M = j;
                                paEngineItem.D1 = (int)(pllInputFreq / pdfFreq);
                                paEngineItem.TargetFreq = targetValue;
                                return true;
                            }
                        }

                    }
                }
            }
            return false;
        }

        private bool IsOkFrc(double tryValue, List<double> targetList, out List<PaEngineItem> paEngineItems)
        {
            bool suitable = true;
            paEngineItems = [];
            foreach (double target in targetList)
            {
                if (TrySingleValue(target, tryValue, out PaEngineItem engine))
                {
                    paEngineItems.Add(engine);
                }
                else
                {
                    suitable = false;
                }
            }
            return suitable;
        }

        public SbcSolutionFrc SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> sbcFreqCalculatorFrcList)
        {
            sbcFreqCalculatorFrcList = [];
            int solutionGroupNo = 0;
            double firstTarget = TargetFreq[0];
            var tryList = TargetFreq.ToList();
            tryList.Remove(TargetFreq[0]);
            var existData = new List<double>();
            existData.Clear();

            for (int hSpeedMode = 1; hSpeedMode <= 2; hSpeedMode++)
            {
                if (ShouldSkipHSpeedMode(firstTarget, hSpeedMode))
                {
                    continue;
                }

                double patGenFreq = hSpeedMode == 1 ? firstTarget : firstTarget / 2;
                if (patGenFreq >= _patGenFreqHighLimit)
                {
                    continue;
                }

                int m2 = hSpeedMode;
                for (int x2 = 1; x2 <= 2; x2++)
                {
                    if (x2 == 1)
                    {
                        int clkD8LowRange = (int)(_clkD8FreqLowLimit / patGenFreq);
                        int clkD8HighRange = (int)(_clkD8FreqHighLimit / patGenFreq);
                        for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                        {
                            double clkD8Freq = patGenFreq * i;
                            TryClkD8FreqForFrc(patGenFreq, clkD8Freq, i, m2, x2, firstTarget, tryList, existData, sbcFreqCalculatorFrcList, ref solutionGroupNo);
                        }
                    }
                    else if (x2 == 2) // D2 only can put 1
                    {
                        double clkD8Freq = patGenFreq / 2;
                        TryClkD8FreqForFrc(patGenFreq, clkD8Freq, 1, m2, x2, firstTarget, tryList, existData, sbcFreqCalculatorFrcList, ref solutionGroupNo);
                    }
                }
            }

            if (sbcFreqCalculatorFrcList.Count > 0)
            {
                return sbcFreqCalculatorFrcList["0"];
            }
            else
            {
                var solutionNull = new SbcSolutionFrc { SbcFreq = 62500000 };
                // 32Hz to 62.5MHz.
                return solutionNull;
            }
        }

        private static bool ShouldSkipHSpeedMode(double firstTarget, int hSpeedMode)
        {
            return (firstTarget <= 250000000 && hSpeedMode == 2)
                || (firstTarget >= 550000000 && hSpeedMode == 1);
        }

        // Shared by the x2==1 (D2 swept) and x2==2 (D2 fixed at 1) branches of SolveSbcFreqFrc:
        // given a candidate clkD8Freq/D2 pair, sweeps the PDF and PLL-input ranges and records
        // any solution IsOkFrc accepts.
        private void TryClkD8FreqForFrc(double patGenFreq, double clkD8Freq, int d2, int m2, int x2, double firstTarget, List<double> tryList, List<double> existData, Dictionary<string, SbcSolutionFrc> sbcFreqCalculatorFrcList, ref int solutionGroupNo)
        {
            if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
            {
                return;
            }

            int pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
            int pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
            for (int j = pdfLowRange; j <= pdfHighRange; j++)
            {
                double pdfFreq = clkD8Freq / j;
                if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                {
                    continue;
                }

                int pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                int pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                {
                    double pllInputFreq = pdfFreq * k;
                    if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                    {
                        continue;
                    }

                    if (!IsOkFrc(pllInputFreq, tryList, out List<PaEngineItem> engines))
                    {
                        continue;
                    }

                    if (existData.Contains(pllInputFreq))
                    {
                        continue;
                    }
                    existData.Add(pllInputFreq);

                    var solution = new SbcSolutionFrc();
                    solution.EngineList.AddRange(engines);
                    solution.SbcFreq = pllInputFreq * 4;
                    var engine = new PaEngineItem
                    {
                        ClkD8Freq = clkD8Freq,
                        M2 = m2,
                        PatGenFreq = patGenFreq,
                        D2 = d2,
                        X2 = x2,
                        PdfFreq = pdfFreq,
                        M = j,
                        PllInputFreq = pllInputFreq,
                        D1 = k,
                        TargetFreq = firstTarget
                    };
                    solution.EngineList.Add(engine);

                    sbcFreqCalculatorFrcList.Add(solutionGroupNo.ToString(), solution);
                    solutionGroupNo++;
                }
            }
        }
    }
}
