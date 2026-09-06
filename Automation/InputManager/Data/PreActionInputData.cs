using System.Collections.Generic;

using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.InputManager.Data
{
    public class PreActionInputData : InputDataBase
    {
        public PinMapSheet PinMapSheet { get; set; }
        public ExcelWorksheet PinGroupSheet { get; set; }
        public List<ChannelMapSheet> ChannelMapSheets { get; set; } = new List<ChannelMapSheet>();
        public IoContiSheet IoContinuity { get; set; } = new IoContiSheet();
    }
}
