using System;
using System.Diagnostics;

namespace BinCutScriptLib.Base
{
    [DebuggerDisplay("{EqName} : {ProductValue} : {Lvcc}")]
    public class PowerStep : IComparable
    {
        public int Bin;
        public double Id;
        //EQ改為記錄整數, eg: read E1,E2..from VddBinDef, parse to 1,2... 
        public int EqName;
        public string EqnBin = string.Empty;
        public double C;
        public double M;
        public double IdsMax;
        public double CpVMax;
        public double CpVMin;
        public double Cp1Gb;
        public double NonCp1Gb;
        public double CpHv;
        public double Lvcc;
        public double ProductValue;
        public string SramthreshCp1 = "";
        public string SramthreshProduct = "";
        public double MonotonicityOffset;
        public double BinningProduct;
        // Don't edit
        public bool IsInterpolationDelete;
        // Don't edit
        public bool IdsOut;
        public bool IsLastPatternPassForCof;
        public bool IsPatternFail;

        public PowerStep() { }

        public PowerStep(PowerStep powerStep)
        {
            if (powerStep == null)
            {
                return;
            }

            Bin = powerStep.Bin;
            Id = powerStep.Id;
            EqName = powerStep.EqName;
            EqnBin = powerStep.EqnBin;
            C = powerStep.C;
            M = powerStep.M;
            IdsMax = powerStep.IdsMax;
            CpVMax = powerStep.CpVMax;
            CpVMin = powerStep.CpVMin;
            Cp1Gb = powerStep.Cp1Gb;
            NonCp1Gb = powerStep.NonCp1Gb;
            CpHv = powerStep.CpHv;
            Lvcc = powerStep.Lvcc;
            ProductValue = powerStep.ProductValue;
            SramthreshCp1 = powerStep.SramthreshCp1;
            SramthreshProduct = powerStep.SramthreshProduct;
            MonotonicityOffset = powerStep.MonotonicityOffset;
            BinningProduct = powerStep.BinningProduct;
            IsInterpolationDelete = powerStep.IsInterpolationDelete;
            IdsOut = powerStep.IdsOut;
            IsLastPatternPassForCof = powerStep.IsLastPatternPassForCof;
            IsPatternFail = powerStep.IsPatternFail;
        }

        public PowerStep Copy()
        {
            return new PowerStep(this);
        }

        public int CompareTo(object? obj)
        {
            if (obj == null)
            {
                return 1;
            }

            var cmpStp = (PowerStep)obj;
            if (Bin > cmpStp.Bin)
            {
                return 1;
            }

            if (Bin < cmpStp.Bin)
            {
                return -1;
            }

            if (EqName > cmpStp.EqName)
            {
                return -1;
            }

            if (EqName < cmpStp.EqName)
            {
                return 1;
            }

            return 0;
        }
    }
}
