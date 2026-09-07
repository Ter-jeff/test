using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData;
using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;

using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Automation.GenerateIgxl.Basic.Business
{
    public class BasicMidResult
    {
        public GlobalSpecSheet GlbSpecSheet { get; set; }

        public List<DcSpecSheet> MultiDcSpecSheets { get; set; }

        public AcInputSheet AcInputSheet { get; set; }

        public AcSpecSheet AcSpecSheet { get; set; }

        public MultiLevelSheets MultiLevelSheets { get; set; }

        public TimeSetSheets TimeSetSheets { get; set; }

        public PatSetSheet PatSetAll { get; set; }

        public PatSetSubSheet PatSetSub { get; set; }
    }
}
