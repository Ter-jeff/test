using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.GenReferenceSheet
{
    public class GenReferenceSheetMain
    {
        private readonly Dictionary<string, int> _referencePriorityList = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "scrrun.dll", 0 },
            { "3", 1 },
            { "CoreTestLibrary.dll", 2 },
            { "Teradyne.TestAPI.dll", 3 },
            { "PlatformInterfacesLibrary.dll", 4 },
            { "PlatformCommonTestLibrary.dll", 5 },
            { "PowerBinning.dll", 6}
        };

        public void GenReferenceSheet()
        {
            ReferenceSheet referenceSheet = TestProgram.IgxlWorkBk.ReferenceSheets.FirstOrDefault().Value;
            if (referenceSheet == null)
            {
                referenceSheet = new ReferenceSheet("References");
                TestProgram.IgxlWorkBk.AddReferenceSheet(FolderStructure.DirReference, referenceSheet);
            }
            ExcelWorksheet referenceBasicConfig = SettingStatic.BasicConfigWorkbook.Worksheets["References"];
            if (referenceBasicConfig != null)
            {
                List<ReferenceRow> referenceRows = new ReadReferenceSheet().ReadSheet(referenceBasicConfig).Rows;
                referenceRows.ForEach(x => referenceSheet.Rows.Add(x));
            }
            var afterOrder = referenceSheet.Rows.OrderBy(x =>
            {
                string fileName = Path.GetFileName(x.FilePath);
                KeyValuePair<string, int> target = _referencePriorityList.FirstOrDefault(y => fileName.ToLower().StartsWith(y.Key.ToLower()));
                if (target.Key != null)
                {
                    return target.Value;
                }

                return 999;
            }).ToList();
            referenceSheet.Rows = afterOrder;
        }
    }
}
