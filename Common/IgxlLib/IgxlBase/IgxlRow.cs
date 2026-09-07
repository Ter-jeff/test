namespace IgxlLib.IgxlBase
{
    public class IgxlRow
    {
        public bool IsBackup { get; set; }
        public string? ColumnA { get; set; }
        public string? SheetName { get; set; }
        public int RowNum { get; set; }

        public IgxlRow()
        {
        }

        protected IgxlRow(IgxlRow igxlRow)
        {
            if (igxlRow == null)
            {
                return;
            }

            SheetName = igxlRow.SheetName;
            RowNum = igxlRow.RowNum;
            IsBackup = igxlRow.IsBackup;
            ColumnA = igxlRow.ColumnA;
        }
    }
}
