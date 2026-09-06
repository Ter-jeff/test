using System.Collections.Generic;

using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.BassData
{
    public class LevelData
    {
        public string LevelSheetName { get; }
        public List<IoInfoRow> IoInfoRows { get; }
        public string DcParameterSyntax { get; set; } = "";
        public string SplitDcBlockSyntax { get; set; } = "";
        public string GlbSymbolSuffix { get; set; } = "";
        public bool CustomInheritance { get; set; } = false;

        public LevelData(string levelSheetName, List<IoInfoRow> ioInfoRows)
        {
            LevelSheetName = levelSheetName;
            IoInfoRows = ioInfoRows;
        }
    }
}
