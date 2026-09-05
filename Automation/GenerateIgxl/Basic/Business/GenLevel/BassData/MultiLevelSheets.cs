using System;
using System.Collections.Generic;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.BassData
{
    public class MultiLevelSheets : Dictionary<string, LevelSheet>
    {
        public MultiLevelSheets() : base(StringComparer.OrdinalIgnoreCase)
        {
        }

        #region Member Function

        public void AddLeveSheet(LevelSheet levelSheet)
        {
            this[levelSheet.Name] = levelSheet;
        }

        #endregion
    }
}
