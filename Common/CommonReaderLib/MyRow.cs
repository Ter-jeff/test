namespace CommonReaderLib
{
    public abstract class MyRow
    {
        public string SheetName { get; set; } = string.Empty;
        public int RowNum { get; set; }

        public string SetColumnA()
        {
            return SheetName + ":Row" + RowNum;
        }

        protected MyRow()
        {
        }

        protected MyRow(MyRow myRow)
        {
            if (myRow == null)
            {
                return;
            }

            SheetName = myRow.SheetName;
            RowNum = myRow.RowNum;
        }
    }
}
