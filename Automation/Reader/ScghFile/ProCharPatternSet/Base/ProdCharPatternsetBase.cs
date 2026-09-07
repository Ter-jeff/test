using System.Collections.Generic;
using System.Linq;

using IgxlLib.IgxlBase;

using ScghLib.Base;

namespace Automation.Reader.ScghFile.ProCharPatternSet.Base
{
    public class PatternWithMode
    {
        public string PatternName = "";
        public string Mode = "";
    }

    public class ProdCharPatternSetBase
    {
        #region result
        public InstanceRow InstanceRow;
        #endregion

        public string SheetName { set; get; }
        public int RowNum { set; get; }
        public string Prefix { set; get; }
        public string InitPatSetNameByNamingRule { set; get; }                                //User definition by naming rule (ex.SocMbist_MS001_MS001_MS002)
        public string PatSetName { set; get; }                                                //After naming rule (Get 0,1,2 segment,  PP_OTCA0_S_PLXX_BI_GR01_BIR_JTG_UNS_ALLFV_SI_SRVA => PP_OTCA0_S)
        public string PayLoadName
        {
            get;
            private set;
        } //Before naming rule (ex. PP_OTCA0_S_PLXX_BI_GR01_BIR_JTG_UNS_ALLFV_SI_SRVA)
        public string InstanceName { set; get; }                                              //After naming rule
        public PatSet PatSet { set; get; }                                                    //ex. PP_OTCA0_S_PLXX_BI_GR01_BIR_JTG_UNS_ALLFV_SI_SRVA

        public List<PatternWithMode> PayloadList
        {
            set
            {
                if (value.Count > 1)
                {
                    var patSet = new PatSet();
                    PayLoadName = PatSet.GetPatSetName(value.Select(x => x.PatternName).ToList());
                }
                else if (value.Count == 1)
                {
                    PayLoadName = value[0].PatternName;
                }
                else
                {
                    PayLoadName = "";
                }

                _payloadList = value;
            }
            get { return _payloadList; }
        }
        public List<string> PayloadAliasList { set; get; }
        public Dictionary<int, PatternWithMode> InitList { set; get; }
        public List<string> InitAliasList { set; get; }
        public bool InitPatternMissing { set; get; }
        public IProdCharItem ProdCharRow { set; get; }
        public bool Nop { set; get; }

        public string Chiplet
        {
            get { return ProdCharRow.Chiplet; }
        }

        private List<PatternWithMode> _payloadList;

        public ProdCharPatternSetBase()
        {
            ProdCharRow = null;
        }

        public ProdCharPatternSetBase(IProdCharItem prodCharRow)
        {
            ProdCharRow = prodCharRow;
            InitList = new Dictionary<int, PatternWithMode>();
            PayloadList = new List<PatternWithMode>();
            InitAliasList = new List<string>();
            PayloadAliasList = new List<string>();
        }
    }
}
