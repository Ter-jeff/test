using System.Collections.Generic;

namespace RfLib.InstrumentSetup.InstrumentTypeData
{
    public class ReplaceInstrumentClass(string instrumenttype, string parameter, Dictionary<string, string> replacekey)
    {
        #region Field
        #endregion

        #region Properity
        public string InstrumentType { set; get; } = instrumenttype;
        public string Parameter { set; get; } = parameter;
        public Dictionary<string, string> ReplaceKey { set; get; } = replacekey;

        #endregion
        #region Constructor
        #endregion
    }
}
