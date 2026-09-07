using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class JobListSheet : IgxlSheet<JobRow>
    {
        public override string SheetType => "DTJobListSheet";
        public override string IgxlSheetName => IgxlSheetNames.JobList;

        public JobListSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public JobListSheet(string sheetName) : base(sheetName)
        {
        }

        public override void Write(string fileName, string version = "")
        {
            Write3P1(fileName, version);
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
            int jobNameIndex = GetIndexFrom(sheetInfo, "Job Name");
            int pinMapIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Pin Map");
            int testInstancesIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Test Instances");
            int flowTableIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Flow Table");
            int aCSpecsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "AC Specs");
            int dCSpecsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "DC Specs");
            int patternSetsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Pattern Sets");
            int binTableIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Bin Table");
            int characterizationIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Characterization");
            int mixedSignalTimingIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Mixed Signal Timing");
            int waveDefinitionsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Wave Definitions");
            int pSetsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Psets");
            int signalsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Signals");
            int portMapIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Port Map");
            int fractionalBusIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Fractional Bus");
            int concurrentSequenceIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Concurrent Sequence");
            int spikeCheckConfigIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "SpikeCheck Config");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                if (sheetInfo.Columns.Column != null)
                {
                    foreach (Column item in sheetInfo.Columns.Column)
                    {
                        if (item.rowIndex == i)
                        {
                            arr[item.indexFrom] = item.columnName;
                        }

                        if (item.Column1 != null)
                        {
                            foreach (Column column1 in item.Column1)
                            {
                                if (column1.rowIndex == i)
                                {
                                    if (version == "3.1")
                                    {
                                        arr[column1.indexFrom] = column1.columnName.Replace("AC Specs", "AC Spec");
                                    }
                                    else
                                    {
                                        arr[column1.indexFrom] = column1.columnName;
                                    }
                                }
                            }
                        }
                    }
                }

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            for (int index = 0; index < Rows.Count; index++)
            {
                JobRow row = Rows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.JobName))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[jobNameIndex] = row.JobName;
                    arr[pinMapIndex] = row.PinMap;
                    arr[testInstancesIndex] = row.TestInstances;
                    arr[flowTableIndex] = row.FlowTable;
                    arr[aCSpecsIndex] = row.AcSpecs;
                    arr[dCSpecsIndex] = row.DcSpecs;
                    arr[patternSetsIndex] = row.PatternSets;
                    arr[binTableIndex] = row.BinTable;
                    arr[characterizationIndex] = row.Characterization;
                    arr[mixedSignalTimingIndex] = row.MixedSignalTiming;
                    arr[waveDefinitionsIndex] = row.WaveDefinitions;
                    arr[pSetsIndex] = row.PSets;
                    arr[signalsIndex] = row.Signals;
                    arr[portMapIndex] = row.PortMap;
                    arr[fractionalBusIndex] = row.FractionalBus;
                    arr[concurrentSequenceIndex] = row.ConcurrentSequence;
                    if (spikeCheckConfigIndex != -1)
                    {
                        arr[spikeCheckConfigIndex] = row.SpikeCheckConfig;
                    }

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

        public JobRow? GetRow(string job)
        {
            return Rows.Find(x => x.JobName.EqualsIgnoreCase(job));
        }
    }
}
