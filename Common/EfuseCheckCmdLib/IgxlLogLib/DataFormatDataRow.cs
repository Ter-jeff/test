using EfuseCheckCmdLib.IgxlLogLib.Base;

namespace EfuseCheckCmdLib.IgxlLogLib
{
    public class DataFormatDataRow
    {
        public EDataFormatType DataFormatType { get; set; }
        public string TestInstance { get; set; } = "";
        public long Number { get; set; }
        public int ActiveSite { get; set; }
        public string TestName { get; set; } = "";
        public string Pin { get; set; } = "";
        public string Channel { get; set; } = "";
        public string Low { get; set; } = "";
        public string Measured { get; set; } = "";
        public string High { get; set; } = "";
        public string Force { get; set; } = "";
        public string Loc { get; set; } = "";
        public string Unit { get; set; } = "";
        public string ForceUnit { get; set; } = "";
        public string TestResult { get; set; } = "";
        public string Pattern { get; set; } = "";
        public int ActuallineNumber { get; set; }

        public string TestType
        {
            get
            {
                if (DataFormatType == EDataFormatType.Measurment)
                {
                    return "Measurement";
                }

                if (DataFormatType == EDataFormatType.Functional)
                {
                    return "Functional Test";
                }

                if (DataFormatType == EDataFormatType.Shmoo)
                {
                    return "Shmoo";
                }

                return "N/A";
            }
        }

    }
}
