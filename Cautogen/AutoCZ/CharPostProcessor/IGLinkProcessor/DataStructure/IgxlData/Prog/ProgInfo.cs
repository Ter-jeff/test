using System.Collections.Generic;

using IgxlLib.IgxlBase;

//using Teradyne.Oasis.IGData;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.IgxlData.Prog
{
    public class ProgInfo
    {
        // data members and fields
        public List<FlowRow> AllFlowSteps;

        public Dictionary<string, string> PinDic;
        public Dictionary<string, PinGroup> PinGroupDic;
        public Dictionary<string, string> PinTypeInChannelDic;

        // constructor
        public ProgInfo()
        {
            AllFlowSteps = new List<FlowRow>();
            PinDic = new Dictionary<string, string>();
            PinGroupDic = new Dictionary<string, PinGroup>();
            PinTypeInChannelDic = new Dictionary<string, string>();
        }
    }
}
