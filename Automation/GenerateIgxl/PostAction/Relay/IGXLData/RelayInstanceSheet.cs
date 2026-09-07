using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.Relay.IGXLData
{
    public class RelayInstanceSheet : InstanceSheet
    {
        public RelayInstanceSheet(string sheetName)
            : base(sheetName)
        {
        }

        public string FolderName;
    }
}
