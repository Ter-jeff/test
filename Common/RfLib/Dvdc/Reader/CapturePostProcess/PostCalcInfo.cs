namespace RfLib.Dvdc.Reader.CapturePostProcess
{
    public class PostCalcInfo
    {
        public string CalcEquation { get; set; }
        public string CalcTestName { get; set; }
        public string CalcStoreName { get; set; }
        public string LowLimit { get; set; }
        public string HiLimit { get; set; }
        public int Bit { get; set; }

        public PostCalcInfo()
        {
            CalcEquation = "";
            CalcTestName = "";
            CalcStoreName = "";
            LowLimit = "";
            HiLimit = "";
            Bit = 0;
        }

        public PostCalcInfo(string calcequation, string calctestname, string calcstorename, string lowlimit, string hilimit, int bitnum)
        {
            CalcEquation = calcequation;
            CalcTestName = calctestname;
            CalcStoreName = calcstorename;
            LowLimit = lowlimit;
            HiLimit = hilimit;
            Bit = bitnum;
        }
    }
}
