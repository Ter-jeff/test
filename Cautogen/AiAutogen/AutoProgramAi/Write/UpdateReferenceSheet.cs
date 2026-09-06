using System;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateReferenceSheet
    {
        public ReferenceSheet Work(ReferenceSheet referenceSheet)
        {
            if (referenceSheet == null)
            {
                referenceSheet = new ReferenceSheet("Reference");
            }
            foreach (ReferenceRow item in referenceSheet.Rows)
            {
                int binIndex = item.FilePath.IndexOf(@"\bin\", StringComparison.OrdinalIgnoreCase);

                if (binIndex < 0)
                {
                    continue;
                }

                item.FilePath = $@".\central_library_cs{item.FilePath.Substring(binIndex)}";
            }

            return referenceSheet;
        }

    }
}
