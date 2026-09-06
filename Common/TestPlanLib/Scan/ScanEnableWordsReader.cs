using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Scan
{
    public class ScanEnableWordsReader : MySheetReader<ScanEnableWordSheet>
    {
        private const string ConEnableWords = "Enable-Words";
        private const string ConRelevantPortions = "Relevant portions";
        private const string ConFlagName = "FlagName";
        private int _indexEnableWords = -1;
        private int _indexRelevantPortions = -1;
        private int _indexFlagName = -1;
        private readonly Dictionary<int, string> _jobIdx = [];
        public Dictionary<string, Dictionary<string, string>> EnableWordsDic = [];
        public Dictionary<string, string> FlagNameDic = [];

        public override ScanEnableWordSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadJob();

            ScanEnableWordSheet scanEnablewordSheet = ReadSheet();

            return scanEnablewordSheet;
        }

        private void ReadJob()
        {
            for (int i = _indexEnableWords + 1; i < ExcelWorksheet.Dimension.End.Column; i++)
            {
                if (i == _indexRelevantPortions || i == _indexFlagName)
                {
                    continue;
                }

                string value = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                _jobIdx.Add(i, value);
            }
        }

        private ScanEnableWordSheet ReadSheet()
        {
            var status = new Dictionary<string, string>();
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                if (_indexEnableWords != -1)
                {
                    string enablewordname = ExcelWorksheet.GetCellValue(i, _indexEnableWords).Trim();
                    if (string.IsNullOrEmpty(enablewordname))
                    {
                        continue;
                    }

                    foreach (int j in _jobIdx.Keys)
                    {
                        status.Add(_jobIdx[j], ExcelWorksheet.GetCellValue(i, j));
                    }
                    if (EnableWordsDic.TryAdd(enablewordname, status))
                    {
                        if (_indexFlagName != -1)
                        {
                            string flagName = ExcelWorksheet.GetCellValue(i, _indexFlagName);
                            if (!string.IsNullOrEmpty(flagName))
                            {
                                FlagNameDic.Add(enablewordname, flagName);
                            }
                        }
                    }
                    status = [];
                }
            }
            return new ScanEnableWordSheet { EnableWordsDic = EnableWordsDic, SheetName = ExcelWorksheet.Name, FlagNameDic = FlagNameDic };
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConEnableWords))
                {
                    _indexEnableWords = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConFlagName))
                {
                    _indexFlagName = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConRelevantPortions))
                {
                    _indexRelevantPortions = i;
                    continue;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConEnableWords))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
