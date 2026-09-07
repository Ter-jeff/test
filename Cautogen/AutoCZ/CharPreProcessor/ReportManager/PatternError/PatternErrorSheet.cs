using Cautogen.common.ReaderWriter.Writer;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError
{
    public class PatternErrorSheet : IExcelSheetWriter
    {
        public void Write(ExcelWorkbook wb)
        {
            if (PatternErrorCache.PatternErrorRowDict.Values.Count == 0)
            {
                return;
            }

            ExcelWorksheet sh = wb.Worksheets.Add("PatternError");

            wb.Worksheets.MoveToStart("PatternError");

            _WriterHeader(sh);

            sh.Cells[2, 1].LoadFromCollection(PatternErrorCache.PatternErrorRowDict.Values);

            sh.Cells.TryAutoFitColumns();
        }

        private static void _WriterHeader(ExcelWorksheet sh)
        {
            sh.Cells[1, 1].Value = "PatternName";
            sh.Cells[1, 2].Value = "SheetName";
            sh.Cells[1, 3].Value = "WrongMeasCount";
            sh.Cells[1, 4].Value = "WrongMeasOrder";
            sh.Cells[1, 5].Value = "MissingMeasPin";
            sh.Cells[1, 6].Value = "MissingPinSeq";
            sh.Cells[1, 7].Value = "MissingPatternInPatInfo";
            sh.Cells[1, 8].Value = "MissingPatternInPatList";
            sh.Cells[1, 9].Value = "PatternShowDontUseInPatList";
            sh.Cells[1, 10].Value = "FileVersionMissMatchGenericNameInPatList";
            sh.Cells[1, 11].Value = "VersionInPatList";
            sh.Cells[1, 12].Value = "PatternDontExistOnServer";
            sh.Cells[1, 13].Value = "PatternCompileSuccess";
            sh.Cells[1, 14].Value = "LatestVersionInServer";
            sh.Cells[1, 1, 1, sh.Dimension.End.Column].Style.Font.Bold = true;
        }
    }
}
