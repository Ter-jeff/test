using System.Collections.Generic;

using Cautogen.AutoCZ.CharPostProcessor.Utility;
using Cautogen.common.IgxlProgramLib.IgxlProgramParser;

using TestPlanLib.PatternListCsvFile;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure
{
    public class InputParam
    {
        public string CharPlan { get; set; }
        public string PatListFile { get; set; }
        public string CharFile { get; set; }
        public string PatInfoFile { get; set; }
        public string PostSettings { get; set; }
        public string OutputFolder { get; set; }
        public string ProgWorkBookPath { get; set; }
        public string JobName { get; set; }
        public string ChannelMapName { get; set; }
        public string CsLibraryPath { get; set; }
        public int StartTnum { get; set; }
        public bool GenCharNotUse { get; set; }
        public bool GenPatNotUse { get; set; }
        public bool GenFlowProdFlow { get; set; }
        public bool GenTxtOnly { get; set; }
        public bool GenAssignSiteVar { get; set; }
        public bool GenTNum { get; set; }
        public bool GenTmpsOnFlow { get; set; }
        public bool GenPmode { get; set; }
        public bool GenPatSub { get; set; }
        public bool GenReadEcid { get; set; }
        public string PatSetFile { get; set; }
        public string ExportVersion { get; set; }
        public string ProjectName { get; set; }
        public Dictionary<string, TimeSetItem> TimeSetVersionDic { get; set; }
        public string TimeSetFolder { get; set; }
        public string PatternFolder { get; set; }
        public Dictionary<string, SubrPatInfo> HardIpInfoAllDict { get; set; }
        public string EnableWords { get; set; }
        public string FlowTmpsName { get; set; }
        public string TNumStart { get; set; }
        public bool UseNewTChar { get; set; }
        public bool IgnoreHfLimits { get; set; }
        public bool RunPayloadAfterSelsram { get; set; }
        public bool GenCSharp { get; set; }
        public IgxlProgram IgxlProgram = null;
        public InputParam()
        {
            CharPlan = "";
            PatSetFile = "";
            PatListFile = "";
            CharFile = "";
            PatInfoFile = "";
            PostSettings = "";
            OutputFolder = "";
            ProgWorkBookPath = "";
            JobName = "";
            ChannelMapName = "";
            GenCharNotUse = false;
            GenPatNotUse = false;
            GenFlowProdFlow = false;
            GenAssignSiteVar = false;
            GenTxtOnly = false;
            GenTNum = false;
            GenTmpsOnFlow = true;
            GenPmode = true;
            GenPatSub = true;
            GenReadEcid = false;
            ExportVersion = "";
            ProjectName = "";
            TimeSetFolder = "";
            PatternFolder = "";
            TimeSetVersionDic = new Dictionary<string, TimeSetItem>();
            FlowTmpsName = "TMPS";
            TNumStart = "10000";
            UseNewTChar = false;
            IgnoreHfLimits = false;
            RunPayloadAfterSelsram = false;
            GenCSharp = false;
            IgxlProgram = null;
        }
    }
}
