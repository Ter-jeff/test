using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

namespace CommonReaderLib.PatternListCsv
{
    public class PatternListSheet : MySheet
    {
        public int IndexFileVersions = -1;
        public int IndexNumber = -1;
        public int IndexPattern = -1;
        public int IndexTimeSetLatest = -1;
        public int IndexUseNotUse = -1;


        public PatternListSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PatternListRow>();
        }

        public List<PatternListRow> Rows { set; get; }


        public void CheckPatternTimeSet(string timeFolder, HashSet<string> patterns)
        {
            var timeSets = Directory.GetFiles(timeFolder).Select(x => Path.GetFileNameWithoutExtension(x)).ToList();
            foreach (PatternListRow row in Rows)
            {
                if (!row.UseNotUse.EqualsIgnoreCase("USE"))
                {
                    continue;
                }

                string timeSet = Path.Combine(timeFolder, row.TimeSet);
                if (!row.TimeSet.EqualsIgnoreCase("NA") &&
                    !string.IsNullOrEmpty(row.TimeSet) &&
                    !timeSets.Any(x => x.IndexOf(Path.GetFileNameWithoutExtension(row.TimeSet), StringComparison.OrdinalIgnoreCase) != -1))
                {
                    AddError(AutoAiErrorType.E_MissingTimesetFile_01, SheetName, row.RowNum, IndexTimeSetLatest, [timeSet]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_MissingTimesetFile_01, SheetName, row.RowNum, IndexTimeSetLatest, [timeSet]);
                }

                if (!row.PatternVersion.EqualsIgnoreCase("NA") &&
                    !row.PatternVersion.EqualsIgnoreCase("N/A") &&
                    !patterns.Contains(row.PatternVersion, StringExtensions.IgnoreCase))
                {
                    AddError(AutoAiErrorType.E_MissingPatternFile_01, SheetName, row.RowNum, IndexFileVersions, [row.PatternVersion]);
                    ErrorReportManager.AddError(AutoAiErrorType.E_MissingPatternFile_01, SheetName, row.RowNum, IndexFileVersions, [row.PatternVersion]);
                }
            }
        }
    }
}
