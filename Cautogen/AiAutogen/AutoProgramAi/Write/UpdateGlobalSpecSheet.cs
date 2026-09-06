using System.Collections.Generic;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateGlobalSpecSheet
    {
        public GlobalSpecSheet Work(GlobalSpecSheet globalSpecSheet, List<string> pins)
        {
            if (globalSpecSheet == null)
            {
                globalSpecSheet = new GlobalSpecSheet("Global Specs");
            }

            foreach (var pin in pins)
            {
                var row = new GlobalSpec(pin);
                row.Value = "0";
                if (!globalSpecSheet.Rows.Exists(x => x.Symbol.EqualsIgnoreCase(pin)))
                {
                    globalSpecSheet.AddRow(row);
                }
            }

            return globalSpecSheet;
        }
    }
}
