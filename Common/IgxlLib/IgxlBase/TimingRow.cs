namespace IgxlLib.IgxlBase
{
    public class TimingRow : IgxlRow
    {
        public string PinGrpName { get; set; } = string.Empty;
        public string PinGrpClockPeriod { get; set; } = string.Empty;
        public string PinGrpSetup { get; set; } = string.Empty;
        public string DataSrc { get; set; } = string.Empty;
        public string DataFmt { get; set; } = string.Empty;
        public string DriveOn { get; set; } = string.Empty;
        public string DriveData { get; set; } = string.Empty;
        public string DriveReturn { get; set; } = string.Empty;
        public string DriveOff { get; set; } = string.Empty;
        public string CompareMode { get; set; } = string.Empty;
        public string CompareOpen { get; set; } = string.Empty;
        public string CompareClose { get; set; } = string.Empty;
        public string CompareRefOffset { get; set; } = string.Empty;
        public string EdgeMode { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string Name { get; internal set; } = string.Empty;
        public string CyclePeriod { get; internal set; } = string.Empty;

        public TimingRow() { }

        public TimingRow(TimingRow timingRow) : base(timingRow)
        {
            if (timingRow == null)
            {
                return;
            }

            PinGrpName = timingRow.PinGrpName;
            PinGrpClockPeriod = timingRow.PinGrpClockPeriod;
            PinGrpSetup = timingRow.PinGrpSetup;
            DataSrc = timingRow.DataSrc;
            DataFmt = timingRow.DataFmt;
            DriveOn = timingRow.DriveOn;
            DriveData = timingRow.DriveData;
            DriveReturn = timingRow.DriveReturn;
            DriveOff = timingRow.DriveOff;
            CompareMode = timingRow.CompareMode;
            CompareOpen = timingRow.CompareOpen;
            CompareClose = timingRow.CompareClose;
            CompareRefOffset = timingRow.CompareRefOffset;
            EdgeMode = timingRow.EdgeMode;
            Comment = timingRow.Comment;
            Name = timingRow.Name;
            CyclePeriod = timingRow.CyclePeriod;
        }

        public TimingRow Copy()
        {
            return new TimingRow(this);
        }
    }
}
