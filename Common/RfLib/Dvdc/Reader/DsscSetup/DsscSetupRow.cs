using System.Collections.Generic;

using CommonReaderLib;

namespace RfLib.Dvdc.Reader.DsscSetup
{
    public class DsscSetupRow : MyRow
    {
        public string Dsscsetup { set; get; }
        public string Pattern { set; get; }
        public string Digsrceqn { set; get; }
        public string Digsrcreg { set; get; }
        public string Digsrcassignment { set; get; }
        public string Digsrcpin { set; get; }
        public string Digsrcsamplesize { set; get; }
        public string Digcappin { set; get; }
        public string Digcapsamplesize { set; get; }
        public string Cusstrdigcapdata { set; get; }

        public string PatModule { set; get; }
        public bool IsBackup { get; set; }
        #region Constructor
        public DsscSetupRow(string sheetName = "")
        {
            SheetName = sheetName;
            Dsscsetup = "";
            Pattern = "";
            Digsrceqn = "";
            Digsrcreg = "";
            Digsrcassignment = "";
            Digsrcpin = "";
            Digsrcsamplesize = "";
            Digcappin = "";
            Digcapsamplesize = "";
            Cusstrdigcapdata = "";
            PatModule = "";
        }
        #endregion

        public List<string> GetInfos()
        {
            return [Dsscsetup,
                Pattern,
                Digsrceqn,
                Digsrcreg,
                Digsrcassignment,
                Digsrcpin,
                Digsrcsamplesize,
                Digcappin,
                Digcapsamplesize,
                Cusstrdigcapdata,
                PatModule];
        }
    }
}
