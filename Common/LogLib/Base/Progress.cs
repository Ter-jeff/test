using CommonLib.Enums;

namespace LogLib.Base
{
    public class Progress
    {
        public int Percentage { get; set; } = 0;
        public string Result { get; set; } = string.Empty;
        public EnumMessageLevel Level { get; set; } = EnumMessageLevel.General;
        public bool EnableMsg { get; set; } = true;
        public string Status { get; set; } = string.Empty;
    }
}
