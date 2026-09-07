using CommonLib.ErrorReport.Base;

namespace RF_PatternTool
{
    public class RFLogManager
    {
        public static List<RFLogError> Errors = new List<RFLogError>();

        public static void Push(RFLogError err, string filename, int row, string header = "")
        {
            err.LogFile = filename;
            err.Header = header;
            err.Row = row;
            Errors.Add(err);
        }
    }
    public class RFLogError
    {
        public string LogFile { get; set; }
        public string Message { get; set; }
        public int Row { get; set; }

        public string LogAddress = "";
        public string Header { get; set; }
        public ErrorType Type { get; set; }
        public EnumErrorLevel Level { get; set; }

        public void SaveAddress(int row, int column)
        {
            if (string.IsNullOrEmpty(LogAddress))
            {

            }
        }
    }

    public enum ErrorType
    {
        InitNotFound,
        DuplicatePatternNameInside,
        DuplicatePatternNameOutside,
        DuplicateTestNameInside,
        DuplicateTestNameOutside,
        DuplicateStoreNameInside,
        DuplicateStoreNameOutside,
        DuplicateCalcStoreNameInside,
        DuplicateCalcStoreNameOutside,
        DuplicateMeasStoreNameInside,
        DuplicateFullStoreNameInside,
        DuplicateMeasStoreNameOutside,
        FieldUseDiffFuse,
        DataNotUpdateByValue,
        DataUpdateByValue,
        TrimInfoNotComplete,
        HeaderMissing,
        PinMissing,
        TestTypeMissing,
        FuseCategoryNotFound,
        FuseSizeMisMatch,
        PatNameIssue,
        EmptyInit,
        FieldValNotValue,
        LimitNotValue,
        FuseReuse,
        PatNameRule,
        CalcStoreNameTestNameErrorNum,
        CalcArgUninitialize,
        CalcFunMismatchWithPyhtonFile,
        WrongWriteCoverage,
        PatternBitError,
        WrongReadCapBits,
        StroeNameErrorWtihCalc,
        NoPatContent,
        FrequencyOutOfRange
    }
}
