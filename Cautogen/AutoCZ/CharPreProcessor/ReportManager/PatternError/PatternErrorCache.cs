using System.Collections.Generic;

namespace Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError
{
    public static class PatternErrorCache
    {
        public static Dictionary<string, PatternErrorRow> PatternErrorRowDict = new Dictionary<string, PatternErrorRow>();

        public enum MarkType
        {
            WrongMeasCount,
            WrongMeasOrder,
            MissingMeasPin,
            MissingPinSeq,
            MissingPatternInPatInfo,
            PatternDontUseInPatList,
            MissingPatternInPatList,
            PatternFileVersionNotMatch,
            PatternDontExistOnServer,
        }

        public static void Reset()
        {
            PatternErrorRowDict.Clear();
        }

        private static PatternErrorRow _GetRow(string patternName, string sheetName)
        {
            if (PatternErrorRowDict.TryGetValue(patternName, out PatternErrorRow row1))
            {
                return row1;
            }

            var row = new PatternErrorRow(patternName) { SheetName = sheetName };
            PatternErrorRowDict.Add(patternName, row);
            return row;
        }

        public static void MarkLatestVersionInServer(string patternName, string sheetName, string patListVersion, string versionInServer)
        {
            PatternErrorRow row = _GetRow(patternName, sheetName);
            row.VersionInPatList = patListVersion;
            row.LatestVersionInServer = versionInServer;
        }

        public static void Mark(string patternName, string sheetName, MarkType markType, string marker = "V")
        {
            PatternErrorRow row = _GetRow(patternName, sheetName);

            switch (markType)
            {
                case MarkType.MissingPatternInPatInfo:
                    row.MissingPatternInPatInfo = marker;
                    break;

                case MarkType.MissingPatternInPatList:
                    row.MissingPatternInPatList = marker;
                    break;

                case MarkType.PatternDontUseInPatList:
                    row.PatternDontUseInPatList = marker;
                    break;

                case MarkType.PatternDontExistOnServer:
                    row.PatternDontExistOnServer = marker;
                    break;

                case MarkType.PatternFileVersionNotMatch:
                    row.PatternFileVersionNotMatch = marker;
                    break;

                case MarkType.WrongMeasOrder:
                    row.WrongMeasCount = marker;
                    break;

                case MarkType.MissingMeasPin:
                    row.MissingMeasPin = marker;
                    break;

                case MarkType.MissingPinSeq:
                    row.MissingPinSeq = marker;
                    break;

                case MarkType.WrongMeasCount:
                    row.WrongMeasCount = marker;
                    break;
            }
        }
    }
}
