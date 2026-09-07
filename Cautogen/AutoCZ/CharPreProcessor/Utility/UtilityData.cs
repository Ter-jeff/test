using System.Collections.Generic;
using System.Data;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor.Utility
{
    public class UtilityData
    {
        public Dictionary<string, string> DeviceMapping = new Dictionary<string, string>();
        public Dictionary<string, PinInfo> PinList = new Dictionary<string, PinInfo>();
        public Dictionary<string, List<string>> PinGroupList = new Dictionary<string, List<string>>();  // pin list from PinList sheet
        public Dictionary<string, List<string>> PinGroups = new Dictionary<string, List<string>>();  // pin groups from pin map file
        public Dictionary<string, string> PinGroupsType = new Dictionary<string, string>();  // type of pin groups mapping
        public List<string> MissingCategory = new List<string>();  // Record the Categories which appeared in input file but not declared in Char_Input_Def
        public Dictionary<string, string> HTestNames = new Dictionary<string, string>();  // for sorting HLN Test Name
        public List<string> DupCategory = new List<string>();
        public DataTable PowerMergeResult = new DataTable();
        public List<HardIpReference> PatInfoErrorList = new List<HardIpReference>();
        public Dictionary<string, HardIpReference> PatInfoDict = new Dictionary<string, HardIpReference>();
        public ParamData InputParam = new ParamData();
        public List<EmaMappingItem> EmaMappingItems = new List<EmaMappingItem>();
        public List<string> AcCategories = new List<string>();
        public List<string> FrcList = new List<string>();
        public List<string> AcSpecsSymbols = new List<string>();
    }
}
