namespace TestPlanLib.BinCut.Checker
{
    public class VddBinSpreadSheetRow
    {
        #region Field
        public string SourceSheetName = "";
        public int RowNum;
        #endregion

        #region Properity
        public string Domain { set; get; }
        public string Mode { set; get; }
        public int Bin { set; get; }
        public string Eqn { set; get; }
        public double C { set; get; }
        public double M { set; get; }
        public object Ids { set; get; }
        public double Cpidsmax { set; get; }
        public double Ftids { set; get; }
        public string Calclvcc { set; get; } = "";
        public double Cpvmin { set; get; }
        public object Lvcc { set; get; }
        public double Cpvmax { set; get; }
        public double Cpgb { set; get; }
        public string Product { set; get; }
        #endregion

        #region Constructor
        public VddBinSpreadSheetRow()
        {
            Domain = "";
            Mode = "";
            Eqn = "";
            Product = "";
            Ids = "";
            Lvcc = "";
        }

        public VddBinSpreadSheetRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Domain = "";
            Mode = "";
            Eqn = "";
            Calclvcc = "";
            Product = "";
            Ids = "";
            Lvcc = "";
        }
        #endregion
    }
}
