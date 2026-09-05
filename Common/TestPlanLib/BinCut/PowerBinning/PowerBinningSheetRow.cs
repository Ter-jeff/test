using System.Collections.Generic;

namespace TestPlanLib.BinCut.PowerBinning
{
    public class PowerBinningSheetRow
    {
        public string SheetName = "";
        public int RowNum;
        public string BinnedMode = "";
        public string BinVoltage = "";
        public string Ids = "";
        public Dictionary<string, double> FactorDictionary;
        public double PowerValue = 0;
        public double ProductValue = 0;
        public double IdsValue = 0;

        #region Constructor
        public PowerBinningSheetRow()
        {
            FactorDictionary = [];
        }

        public PowerBinningSheetRow(string sheetName)
        {
            SheetName = sheetName;
            FactorDictionary = [];
        }

        public PowerBinningSheetRow(PowerBinningSheetRow powerBinningSheetRow)
        {
            FactorDictionary = [];
            if (powerBinningSheetRow == null)
            {
                return;
            }

            SheetName = powerBinningSheetRow.SheetName;
            RowNum = powerBinningSheetRow.RowNum;
            BinnedMode = powerBinningSheetRow.BinnedMode;
            BinVoltage = powerBinningSheetRow.BinVoltage;
            Ids = powerBinningSheetRow.Ids;
            FactorDictionary = powerBinningSheetRow.FactorDictionary != null ? new Dictionary<string, double>(powerBinningSheetRow.FactorDictionary) : [];
            PowerValue = powerBinningSheetRow.PowerValue;
            ProductValue = powerBinningSheetRow.ProductValue;
            IdsValue = powerBinningSheetRow.IdsValue;
        }
        #endregion

        public PowerBinningSheetRow Copy()
        {
            return new PowerBinningSheetRow(this);
        }
    }
}
