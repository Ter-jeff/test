using System.Collections.Generic;

namespace Automation.PreCheck.AllParaData
{
    public class PostActionParaData : ParaData
    {
        private List<string> _nWireSheetList;
        private List<string> _relaySheetList;

        public bool FlagEcidSortingGen { get; set; } = false;
        public List<string> RelaySheetList
        {
            get
            {
                return _relaySheetList ?? (_relaySheetList = new List<string>());
            }
            set
            {
                _relaySheetList = value;
            }
        }
        public List<string> NWireSheetList
        {
            get
            {
                return _nWireSheetList ?? (_nWireSheetList = new List<string>());
            }
            set
            {
                _nWireSheetList = value;
            }
        }
        public string FreeRunClkTdrTrue32Clk8IdleXml { get; set; }
        public string TestInstCommon { get; set; }
        public bool InitMainFlowFlag { set; get; } = true;
        public string TestNumberXlsx { get; set; }
    }
}
