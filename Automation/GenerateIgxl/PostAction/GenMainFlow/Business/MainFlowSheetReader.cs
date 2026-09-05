using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;

using CommonLib.Utility;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Business
{
    public class MainFlowSheetReader
    {
        private const string ConSequence = "Sequence";
        private const string ConComment = "Comment";
        private ExcelWorksheet _mainFloWorksheet;
        private readonly Dictionary<string, int> _headerColumnDictionary = new Dictionary<string, int>();
        private readonly List<MainFlowBase> _mainFlowBases = new List<MainFlowBase>();

        private int _startCol;
        private int _startRow;
        private int _endCol;
        private int _endRow;

        public MainFlowSheet ReadSheet(ExcelWorksheet pWorksheet)
        {
            _mainFloWorksheet = pWorksheet;

            ReadHeader();

            ReadData();

            MainFlowSheet mainFlowSheet = new MainFlowSheet(_mainFlowBases, _headerColumnDictionary.Keys.ToList());

            return mainFlowSheet;
        }

        private void ReadHeader()
        {
            string header;
            _startCol = _mainFloWorksheet.Dimension.Start.Column;
            _startRow = _mainFloWorksheet.Dimension.Start.Row;
            _endCol = _mainFloWorksheet.Dimension.End.Column;
            _endRow = _mainFloWorksheet.Dimension.End.Row;

            #region Check is Ibiza format

            bool isCheckBoxformat = false;

            for (int i = _startCol; i <= _endCol; i++)
            {
                string first = GetCellValue(1, i);
                if (CellDiff.IsLiked(first, "Flow Name"))
                {
                    isCheckBoxformat = true;
                    break;
                }
            }

            #endregion

            if (isCheckBoxformat)
            {
                #region Is Ibiza format

                for (int i = _startRow + 3; i <= _endRow; i++)
                {
                    header = GetCellValue(1, i);
                    if (CellDiff.IsLiked(header, ConComment))
                    {
                        break;
                    }
                    foreach (string subJob in header.Split(','))
                    {
                        _headerColumnDictionary.Add(subJob.Split('(')[0].Trim(), i);
                    }
                }

                #endregion
            }
            else
            {
                for (int i = _startRow; i <= _endRow; i++)
                {
                    for (int j = _startCol; j < _endCol; j++)
                    {
                        header = GetCellValue(i, j);
                        if (CellDiff.IsLiked(header, ConSequence))
                        {
                            _startCol = j;
                            _startRow = i - 1;
                            break;
                        }
                    }
                }

                for (int i = _startCol; i <= _endCol; i++)
                {
                    header = GetCellValue(_startRow + 1, i).Trim();
                    if (CellDiff.IsLiked(header, ConSequence))
                    {
                        header = GetCellValue(_startRow, i).Split('(')[0].Trim();
                        foreach (string subJob in header.Split(','))
                        {
                            _headerColumnDictionary.Add(subJob.Split('(')[0].Trim(), i);
                        }
                    }
                }
            }

        }

        private void ReadData()
        {
            #region Check is Ibiza format
            bool isCheckBoxformat = false;
            for (int i = _startCol; i <= _endCol; i++)
            {
                string first = GetCellValue(1, i);
                if (CellDiff.IsLiked(first, "Flow Name"))
                {
                    isCheckBoxformat = true;
                    break;
                }
            }

            #endregion

            if (isCheckBoxformat)
            {
                foreach (KeyValuePair<string, int> keyValuePair in _headerColumnDictionary)
                {
                    string header = keyValuePair.Key;
                    int column = keyValuePair.Value;
                    MainFlowBase mainFlowBase = new MainFlowBase { JobName = header.Replace(" ", "_") };

                    for (int i = _startRow + 1; i <= _endRow; i++)
                    {
                        string content = GetCellValue(i, column).Trim();
                        if (CellDiff.IsLiked(content, "V"))
                        {
                            content = GetCellValue(i, 2).Trim();
                        }
                        else
                        {
                            continue;
                        }

                        FlowSequence flowName = new FlowSequence(content) { Job = mainFlowBase.JobName, RowNum = i };
                        mainFlowBase.Sequences.Add(flowName);
                    }
                    mainFlowBase.MainFlowName = "Main_Flow_" + mainFlowBase.JobName;
                    _mainFlowBases.Add(mainFlowBase);
                }
            }
            else
            {
                foreach (KeyValuePair<string, int> keyValuePair in _headerColumnDictionary)
                {
                    string header = keyValuePair.Key;
                    int column = keyValuePair.Value;
                    MainFlowBase mainFlowBase = new MainFlowBase { JobName = header.Replace(" ", "_") };
                    for (int i = _startRow + 2; i <= _endRow; i++)
                    {
                        string content = GetCellValue(i, column).Trim();
                        FlowSequence flowName = new FlowSequence(content) { Job = mainFlowBase.JobName, RowNum = i };
                        if (content != "")
                        {
                            mainFlowBase.Sequences.Add(flowName);
                        }
                    }
                    mainFlowBase.MainFlowName = "Main_Flow_" + mainFlowBase.JobName;
                    _mainFlowBases.Add(mainFlowBase);
                }
            }
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = _mainFloWorksheet.Cells[rowNumber, columnNumber].Value;
            if (value != null)
            {
                return value.ToString();
            }
            return "";
        }
    }
}
