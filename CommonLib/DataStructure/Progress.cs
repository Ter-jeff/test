namespace CommonLib.DataStructure
{
    public class ProgressStatus
    {
        public int Percentage { get; set; } = 0;
        public string Result { get; set; } = string.Empty;
        public MessageLevel Level { get; set; } = MessageLevel.General;
        public bool EnableMsg { get; set; } = true;
        public string Status { get; set; } = string.Empty;
    }
}
