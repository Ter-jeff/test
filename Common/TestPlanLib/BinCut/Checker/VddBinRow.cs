using System.Collections.Generic;

namespace TestPlanLib.BinCut.Checker
{
    public class VddBinRow
    {
        public bool IsOtherRail;
        public string SheetName = "";
        public int TableIdx;
        public int RowNum;
        public string Domain = "";
        public string Mode = "";
        public int Bin;
        public string Id = "";
        public string Eq = "";
        public int Eqn;
        public double C;
        public double M;
        public double CpIdsMax;
        public double FtIds;
        public double CpVMax;
        public double CpVMin;
        public double CpGb;
        public string AllowEqual = "";
        public double MonoDelta = 0.0;
        public List<double> LvccValues = [];
        public List<double> ProductValues = [];
        public List<double> LvccValuesWithMono = [];
        public List<double> ProductValuesWithMono = [];
    }
}
