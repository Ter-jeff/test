using System.Collections.Generic;
using System.Linq;

namespace Automation.Utility
{
    public class SbcFreqCalculator
    {
        private readonly int _clkD8FreqLowLimit = 125000000;
        private readonly int _clkD8FreqHighLimit = 275000000;
        private readonly int _pdfFreqLowLimit = 3000000;
        private readonly int _pdfFreqHighLimit = 6000000;
        private readonly int _pllInputFreqHighLimit = 200000000;
        private readonly int _pllInputFreqLowLimit = 3000000;

        public string Sdf { set; get; }
        public List<int> TargetFreq = new List<int>();

        public SbcSolution SolveSbcFreq()
        {
            SbcSolution solutions = new SbcSolution();

            if (!TargetFreq.Any())
            {
                return solutions;  //if list is 0,  Max function  
            }

            int maxTarget = TargetFreq.Max();
            List<int> tryList = TargetFreq.ToList();
            tryList.Remove(maxTarget);
            int clkD8LowRange = _clkD8FreqLowLimit / maxTarget;
            int clkD8HighRange = _clkD8FreqHighLimit / maxTarget;
            for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
            {
                int clkD8Freq = maxTarget * i;
                if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                {
                    continue;
                }
                int pdfHighRange = clkD8Freq / _pdfFreqLowLimit;
                int pdfLowRange = clkD8Freq / _pdfFreqHighLimit;

                for (int j = pdfLowRange; j <= pdfHighRange; j++)
                {
                    int pdfFreq = clkD8Freq / j;

                    if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                    {
                        continue;
                    }
                    int pllInputLowRange = _pllInputFreqLowLimit / pdfFreq;
                    int pllInputHighRange = _pllInputFreqHighLimit / pdfFreq;
                    for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                    {
                        int pllInputFreq = pdfFreq * k;
                        if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                        {
                            continue;
                        }

                        if (IsOk(pllInputFreq, tryList, out List<PaEngine> engines))
                        {
                            SbcSolution solution = new SbcSolution();
                            solution.EngineList.AddRange(engines);
                            solution.SbcFreq = pllInputFreq * 4;
                            PaEngine engine = new PaEngine
                            {
                                ClkD8Freq = clkD8Freq,
                                D2 = i,
                                PdfFreq = pdfFreq,
                                M = j,
                                PllInputFreq = pllInputFreq,
                                D1 = k,
                                TargetFreq = maxTarget
                            };
                            solution.EngineList.Add(engine);
                            if (solution.SbcFreq < 62500000)     // 32Hz to 62.5MHz.
                            {
                                return solution;
                            }
                        }
                    }
                }
            }
            SbcSolution solutionNull = new SbcSolution { SbcFreq = 62500000 };
            return solutionNull;
        }

        internal bool IsOk(int tryValue, List<int> targetList, out List<PaEngine> engines)
        {
            bool suitable = true;
            engines = new List<PaEngine>();
            foreach (int target in targetList)
            {
                if (TrySingelValue(target, tryValue, out PaEngine engine))
                {
                    engines.Add(engine);
                }
                else
                {
                    suitable = false;
                }
            }
            return suitable;
        }

        internal bool TrySingelValue(int targetValue, int pllInputFreq, out PaEngine engine)
        {
            if (pllInputFreq.Equals(7680000))
            {
            }
            engine = new PaEngine();
            int clkD8Freq, pdfFreq;
            int pdfLowRange, pdfHighRange;
            int clkD8LowRange = _clkD8FreqLowLimit / targetValue;
            int clkD8HighRange = _clkD8FreqHighLimit / targetValue;
            for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
            {
                clkD8Freq = targetValue * i;
                if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                {
                    continue;
                }
                pdfHighRange = clkD8Freq / _pdfFreqLowLimit;
                pdfLowRange = clkD8Freq / _pdfFreqHighLimit;
                if (clkD8Freq.Equals(245760000))
                {
                }
                for (int j = pdfLowRange; j <= pdfHighRange; j++)
                {

                    pdfFreq = clkD8Freq / j;
                    if (pdfFreq.Equals(3840000))
                    {
                    }
                    if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                    {
                        continue;
                    }
                    if (pllInputFreq % pdfFreq == 0)
                    {
                        engine.ClkD8Freq = clkD8Freq;
                        engine.D2 = i;
                        engine.PdfFreq = pdfFreq;
                        engine.PllInputFreq = pllInputFreq;
                        engine.M = j;
                        engine.D1 = pllInputFreq / pdfFreq;
                        engine.TargetFreq = targetValue;
                        return true;
                    }
                }
            }
            return false;
        }
    }

    public class PaEngine
    {
        public int TargetFreq { set; get; }
        public int D2 { set; get; }
        public int ClkD8Freq { set; get; }
        public int M { set; get; }
        public int PdfFreq { set; get; }
        public int D1 { set; get; }
        public int PllInputFreq { set; get; }
    }

    public class SbcSolution
    {
        public int SbcFreq { set; get; }
        public List<PaEngine> EngineList { set; get; } = new List<PaEngine>();
    }
}
