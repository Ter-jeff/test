using System.Text.RegularExpressions;

namespace TestPlanLib.DataStruct
{
    public partial class IoContiRow
    {
        public const string ConHeaderGroup = "Group";
        public const string ConHeaderDual = "Dual row location";
        public const string ConHeaderBumpName = "Bump Name";
        public const string ConHeaderBumpNumber = "Bump Number";
        public const string ConHeaderBallName = "Ball Name";
        public const string ConHeaderSocBall = "SoC Ball Location";
        public const string ConHeaderPadName = "Pad Name";
        public const string ConHeaderSignal = "Signal Name";
        public const string ConHeaderDir = "Dir";
        public const string ConHeaderCellName = "Cell Name";
        public const string ConHeaderPadGroup = "Pad Group";
        public const string ConHeaderIoVoltage = "I/O Voltage";
        public const string ConHeaderFsDd = "FS/DD";
        public const string ConHeaderChiplet = "Chiplet";

        [GeneratedRegex(@"(?<voltage>\d+[.]\d+)[0]+")]
        private static partial Regex MyRegex();

        public string Group { set; get; } = "";
        public string DualRowLocation { set; get; } = "";
        public string BumpName { set; get; } = "";
        public string BumpNumber { set; get; } = "";
        public string BallName { set; get; } = "";
        public string SocBallLocation { set; get; } = "";
        public string PadName { set; get; } = "";
        public string SignalName { set; get; } = "";
        public string Dir { set; get; } = "";
        public string CellName { set; get; } = "";
        public string PadGroup { set; get; } = "";
        public string IoVoltage { set; get; } = "";
        public string FsDd { set; get; } = "";
        public string Chiplet { set; get; } = "";
        public int RowNum { set; get; } = 0;

        public string GetVoltageGroupName()
        {
            if (IoVoltage.Length == 0)
            {
                return "";
            }

            string pinGrpName;
            string voltage =
                    IoVoltage.Replace("V", "")
                        .Replace("v", "")
                        .Replace("m", "")
                        .Replace(" ", "")
                        .Replace("\t", "");

            if (voltage.Contains('.'))
            {
                pinGrpName = "Pins_" + voltage.Replace(".", "p") + "v";
                string name = pinGrpName;
                pinGrpName = MyRegex().Replace(pinGrpName, delegate
                {
                    string num = MyRegex().Match(name).Groups["voltage"].ToString();
                    return num;
                });
            }
            else
            {
                pinGrpName = "Pins_" + voltage + "p0v";
            }
            if (!string.IsNullOrEmpty(Chiplet))
            {
                pinGrpName = pinGrpName + "_" + Chiplet;
            }

            return pinGrpName;
        }
    }
}
