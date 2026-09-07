using System.IO;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadJobListSheet : IgxlSheetReader<JobListSheet>
    {
        private const string ConHeaderJobName = "Job Name";
        private const string ConHeaderPinMap = "Pin Map";
        private const string ConHeaderTestIns = "Test Instances";
        private const string ConHeaderFlowTable = "Flow Table";
        private const string ConHeaderAcSpecs = "AC Specs";
        private const string ConHeaderAcSpec = "AC Spec";
        private const string ConHeaderDcSpec = "DC Specs";
        private const string ConHeaderPatternSets = "Pattern Sets";
        private const string ConHeaderBinTable = "Bin Table";
        private const string ConHeaderCharacterization = "Characterization";
        private const string ConHeaderMixSignalTiming = "Mixed Signal Timing";
        private const string ConHeaderWaveDef = "Wave Definitions";
        private const string ConHeaderPSets = "Psets";
        private const string ConHeaderSignals = "Signals";
        private const string ConHeaderPortMap = "Port Map";
        private const string ConHeaderFraction = "Fractional Bus";
        private const string ConHeaderConcurrent = "Concurrent Sequence";
        private const string ConHeaderComment = "Comment";

        private const int StartRowIndex = 4;
        private const int StartColumnIndex = 2;

        public override EnumSheetType SupportedType => EnumSheetType.DTJobListSheet;

        private ExcelWorksheet _excelWorksheet = null!;
        private static JobListSheet _jobListSheet = null!;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _jobnameIndex = -1;
        private int _pinMapIndex = -1;
        private int _testInsIndex = -1;
        private int _flowTabIndex = -1;
        private int _acSpecIndex = -1;
        private int _dcSpecIndex = -1;
        private int _patternSetsIndex = -1;
        private int _binTableIndex = -1;
        private int _characterizationIndex = -1;
        private int _mixSignalIndex = -1;
        private int _waveDefineIndex = -1;
        private int _pSetsIndex = -1;
        private int _signalIndex = -1;
        private int _portMapIndex = -1;
        private int _fractionalIndex = -1;
        private int _concurrentSeqIndex = -1;
        private int _commentIndex = -1;

        public override JobListSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null!;
            }

            _excelWorksheet = excelWorksheet;

            _jobListSheet = new JobListSheet(excelWorksheet) { Name = excelWorksheet.Name };

            Reset();

            if (!GetDimensions())
            {
                return null!;
            }

            if (!GetFirstHeaderPosition())
            {
                return null!;
            }

            if (!GetHeaderIndex())
            {
                return null!;
            }

            _jobListSheet = ReadSheetData();

            return _jobListSheet;
        }

        public override JobListSheet ReadSheet(Stream stream, string sheetName)
        {
            var jobListSheet = new JobListSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine()!;
                    if (i > StartRowIndex)
                    {
                        JobRow jobRow = GetJobRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(jobRow.JobName))
                        {
                            isBackup = true;
                            continue;
                        }

                        jobRow.IsBackup = isBackup;
                        jobListSheet.AddRow(jobRow);
                    }

                    i++;
                }
            }

            return jobListSheet;
        }

        private JobRow GetJobRow(string line, string sheetName, int row)
        {
            string[] arr = line.Split('\t');
            var jobRow = new JobRow
            {
                RowNum = row,
                SheetName = sheetName
            };
            int index = StartColumnIndex - 1;
            string content = GetCellValue(arr, 0);
            jobRow.ColumnA = content;
            content = GetCellValue(arr, index);
            jobRow.JobName = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.PinMap = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.TestInstances = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.FlowTable = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.AcSpecs = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.DcSpecs = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.PatternSets = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.BinTable = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.Characterization = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.MixedSignalTiming = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.WaveDefinitions = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.PSets = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.Signals = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.PortMap = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.FractionalBus = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.ConcurrentSequence = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.SpikeCheckConfig = content;
            index++;
            content = GetCellValue(arr, index);
            jobRow.Comment = content;
            return jobRow;
        }

        private JobListSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new JobRow { RowNum = i };
                if (_jobnameIndex != -1)
                {
                    row.JobName = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _jobnameIndex).Trim();
                }

                if (_pinMapIndex != -1)
                {
                    row.PinMap = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _pinMapIndex).Trim();
                }

                if (_testInsIndex != -1)
                {
                    row.TestInstances = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _testInsIndex).Trim();
                }

                if (_flowTabIndex != -1)
                {
                    row.FlowTable = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _flowTabIndex).Trim();
                }

                if (_acSpecIndex != -1)
                {
                    row.AcSpecs = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _acSpecIndex).Trim();
                }

                if (_dcSpecIndex != -1)
                {
                    row.DcSpecs = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _dcSpecIndex).Trim();
                }

                if (_patternSetsIndex != -1)
                {
                    row.PatternSets = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _patternSetsIndex).Trim();
                }

                if (_binTableIndex != -1)
                {
                    row.BinTable = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _binTableIndex).Trim();
                }

                if (_characterizationIndex != -1)
                {
                    row.Characterization = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _characterizationIndex).Trim();
                }

                if (_mixSignalIndex != -1)
                {
                    row.MixedSignalTiming = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _mixSignalIndex).Trim();
                }

                if (_waveDefineIndex != -1)
                {
                    row.WaveDefinitions = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _waveDefineIndex).Trim();
                }

                if (_pSetsIndex != -1)
                {
                    row.PSets = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _pSetsIndex).Trim();
                }

                if (_signalIndex != -1)
                {
                    row.Signals = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _signalIndex).Trim();
                }

                if (_portMapIndex != -1)
                {
                    row.PortMap = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _portMapIndex).Trim();
                }

                if (_fractionalIndex != -1)
                {
                    row.FractionalBus = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _fractionalIndex).Trim();
                }

                if (_concurrentSeqIndex != -1)
                {
                    row.ConcurrentSequence = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _concurrentSeqIndex).Trim();
                }

                if (_commentIndex != -1)
                {
                    row.Comment = EpplusExtensions.GetMergedCellValue(_excelWorksheet, i, _commentIndex).Trim();
                }

                _jobListSheet.AddRow(row);
            }
            return _jobListSheet;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _jobnameIndex = -1;
            _pinMapIndex = -1;
            _testInsIndex = -1;
            _flowTabIndex = -1;
            _acSpecIndex = -1;
            _dcSpecIndex = -1;
            _patternSetsIndex = -1;
            _binTableIndex = -1;
            _characterizationIndex = -1;
            _mixSignalIndex = -1;
            _waveDefineIndex = -1;
            _pSetsIndex = -1;
            _signalIndex = -1;
            _portMapIndex = -1;
            _fractionalIndex = -1;
            _concurrentSeqIndex = -1;
            _commentIndex = -1;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string header = EpplusExtensions.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderJobName))
                {
                    _jobnameIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPinMap))
                {
                    _pinMapIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderTestIns))
                {
                    _testInsIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderFlowTable))
                {
                    _flowTabIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderAcSpec) ||
                     header.EqualsIgnoreCase(ConHeaderAcSpecs))
                {
                    _acSpecIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderDcSpec))
                {
                    _dcSpecIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPatternSets))
                {
                    _patternSetsIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderBinTable))
                {
                    _binTableIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderCharacterization))
                {
                    _characterizationIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderMixSignalTiming))
                {
                    _mixSignalIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderWaveDef))
                {
                    _waveDefineIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPSets))
                {
                    _pSetsIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderSignals))
                {
                    _signalIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPortMap))
                {
                    _portMapIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderFraction))
                {
                    _fractionalIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderConcurrent))
                {
                    _concurrentSeqIndex = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderComment))
                {
                    _commentIndex = i;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet, i, j).Trim().EqualsIgnoreCase(ConHeaderJobName))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }
    }
}
