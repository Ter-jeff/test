using System.Collections.Generic;

namespace RfLib.InstrumentSetup.InstrumentTypeData
{
    public class ReplaceInstrumentSheet
    {
        #region Field
        #endregion

        #region Properity
        public string SheetName { get; set; } = "";
        public List<ReplaceInstrumentClass> ReplaceClass { get; } = [];
        #endregion
    }
}
