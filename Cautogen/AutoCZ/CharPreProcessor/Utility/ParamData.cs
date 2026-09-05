
using System.Collections.Generic;

using Cautogen.common.IgxlProgramLib.IgxlProgramParser;

namespace Cautogen.AutoCZ.CharPreProcessor.Utility
{
    public class ParamData
    {
        public string CharPlanFile { get; set; }
        public List<string> CharPlanFileList { get; set; }
        public string JobName { get; set; }
        public string DefFile { get; set; }
        public string PatinfoFile { get; set; }
        public string PatListFile { get; set; }
        public string PinMapFile { get; set; }
        public string GlobalSpecsFile { get; set; }
        public string AcSpecsFile { get; set; }
        public string SelSramMappingTable { get; set; }
        public string SpiPatSetFile { get; set; }
        public string TarDic { get; set; }
        public string ProjectName { get; set; }
        public bool NvIsChecked { get; set; }
        public bool IsSplitSgmt32 { get; set; }
        public bool IsUseRtosCmd { get; set; }
        public bool IsMergeHlv { get; set; }
        public string BaseProgram { get; set; }
        public bool ShmooPowerPinHightoLow { get; set; }
        public bool CharPreCheckWithoutTp { get; set; }
        public bool CharPreCheckForNewTChar { get; set; }
        public bool CommandLineMode { get; set; }

        public IgxlProgram IgxlProgram = null;

        public ParamData()
        {
            CharPlanFile = "";
            JobName = "";
            DefFile = "";
            PatinfoFile = "";
            PatListFile = "";
            PinMapFile = "";
            GlobalSpecsFile = "";
            AcSpecsFile = "";
            SelSramMappingTable = "";
            SpiPatSetFile = "";
            TarDic = "";
            NvIsChecked = false;
            IsSplitSgmt32 = false;
            IsUseRtosCmd = true;
            IsMergeHlv = true;
            BaseProgram = "";
            ShmooPowerPinHightoLow = false;
            CharPreCheckWithoutTp = false;
            CharPreCheckForNewTChar = true;
            IgxlProgram = null;
            CharPlanFileList = [];
        }
    }
}
