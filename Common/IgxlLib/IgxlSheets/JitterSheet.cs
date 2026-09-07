using System.IO;
using System.Linq;
using System.Text;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class JitterSheet : IgxlSheet<JitterRow>
    {
        public override string SheetType => "DTSerialJitterSheet";
        public override string IgxlSheetName => IgxlSheetNames.JitterSheet;

        public JitterSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public JitterSheet(string sheetName) : base(sheetName)
        {
        }

        protected override void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int jitterSetIndex = GetIndexFrom(sheetInfo, "Jitter Set");
            int pinIndex = GetIndexFrom(sheetInfo, "Pin/Group");
            int modeIndex = GetIndexFrom(sheetInfo, "Insertion", "Mode");
            int periodIndex = GetIndexFrom(sheetInfo, "Insertion", "Period");
            int amplitudeIndex = GetIndexFrom(sheetInfo, "Insertion", "Amplitude in UI");
            int sampleTypeIndex = GetIndexFrom(sheetInfo, "Measurement", "Sample Type");
            int targetResolutionIndex = GetIndexFrom(sheetInfo, "Measurement", "Target Resolution");
            int targetPatternRepsIndex = GetIndexFrom(sheetInfo, "Measurement", "Target Pattern Reps");
            int bitsPerPatternIndex = GetIndexFrom(sheetInfo, "Measurement", "Bits Per Pattern");
            int minPatternSkipsIndex = GetIndexFrom(sheetInfo, "Measurement", "Min Pattern Skips");
            int bitPeriodIndex = GetIndexFrom(sheetInfo, "Measurement", "Bit Period");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                JitterRow row = Rows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.JitterSet))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[jitterSetIndex] = row.JitterSet;
                    arr[pinIndex] = row.PinOrGroup;
                    arr[modeIndex] = row.Mode;
                    arr[periodIndex] = row.Period;
                    arr[amplitudeIndex] = row.AmplitudeInUi;
                    arr[sampleTypeIndex] = row.SampleType;
                    arr[targetResolutionIndex] = row.TargetResolution;
                    arr[targetPatternRepsIndex] = row.TargetPatternReps;
                    arr[bitsPerPatternIndex] = row.BitsPerPattern;
                    arr[minPatternSkipsIndex] = row.MinPatternSkips;
                    arr[bitPeriodIndex] = row.BitPeriod;
                    arr[commentIndex] = row.Comment;
                }
                else
                {
                    arr = ["\t"];
                }
                sb.AppendLine(string.Join("\t", arr));
            }
            #endregion

            sw.Write(sb.ToString());
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P1(fileName, version);
        }
    }
}
