namespace IgxlLib.IgxlBase
{
    public class LevelRow : IgxlRow
    {
        public string? PinName { get; set; }
        public string Seq { get; set; } = string.Empty;
        public string? Parameter { get; set; }
        public string? Value { get; set; }
        public string? Comment { get; set; }

        public LevelRow()
        {
        }

        public LevelRow(string pinName, string parameter, string value, string comment, int rowNum = 0)
        {
            PinName = pinName;
            Parameter = parameter;
            Value = value;
            Comment = comment;
            RowNum = rowNum;
        }

        public LevelRow(LevelRow levelRow) : base(levelRow)
        {
            if (levelRow == null)
            {
                return;
            }

            PinName = levelRow.PinName;
            Seq = levelRow.Seq;
            Parameter = levelRow.Parameter;
            Value = levelRow.Value;
            Comment = levelRow.Comment;
            RowNum = levelRow.RowNum;
        }

        public LevelRow Copy()
        {
            return new LevelRow(this);
        }
    }
}
