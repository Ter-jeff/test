using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.Static;

using CommonLib.Utility;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Business
{
    public class MainFlowSheetReaderNew
    {
        #region Field
        private const string ConSource = "Source";
        private const string ConSheetName1 = "SheetName:SubFlow";
        private const string ConSheetName2 = "SheetName";
        private const string ConSubFlow = "SubFlow";
        private const string ConSubProgram = "SubProgram";
        private const string ConEnable1 = "Enable";
        private const string ConEnable2 = @"^Enable\s*\(all sites\)";
        private const string ConBintableEnable = @"^BinTable Enable\s*\(all sites\)";
        private const string ConSiteFlagPerSite = "SiteFlag (per site)";
        private const string ConFailFlag = "FailFlag";
        private const string ConComment = "Comment";
        private const string ConModule = "Module";
        private const string ConInclude = "Include";
        private const string ConGroup = "Group";
        private const string ConOption = "Option";
        private const string ConOptions = "Options";
        private ExcelWorksheet _excelWorksheet;
        private readonly Dictionary<string, int> _headerColumnDictionary = new Dictionary<string, int>();
        private readonly Dictionary<string, List<int>> _t0txColumnDictionary = new Dictionary<string, List<int>>();
        private readonly List<MainFlowBase> _mainFlowBases = new List<MainFlowBase>();

        private List<string> _enableModules;
        private int _startCol;
        private int _startRow;
        private int _endCol;
        private int _endRow;
        private int _sourceCol = -1;
        private int _sheetNameCol = -1;
        private int _subFlowCol = -1;
        private int _subProgram = -1;
        private int _enableWdCol = -1;
        private int _bintableEnableWdCol = -1;
        private int _siteFlagPerSiteCol = -1;
        private int _failFlagCol = -1;
        private int _commentCol = -1;
        private int _moduleCol = -1;
        private int _includeCol = -1;
        private int _groupCol = -1;
        private int _optionCol = -1;
        private readonly List<int> _columnIndex = new List<int>();
        #endregion

        public List<string> GetEnableModules()
        {
            return _enableModules;
        }

        public MainFlowSheet ReadSheet(ExcelWorksheet pWorksheet)
        {
            _excelWorksheet = pWorksheet;

            ReadHeader();

            ReadDataMain();

            if (LocalSpecs.Options.GenerateT0TXTestprogram)
            {
                ReadT0TXDataMain();
            }

            if (_subProgram != -1)
            {
                ReadDataSub();
            }

            ReadEnableModule();

            MainFlowSheet mainFlowSheet = new MainFlowSheet(
                mainFlowBases: _mainFlowBases,
                sourceCol: _sourceCol,
                sheetNameCol: _sheetNameCol,
                subFlowCol: _subFlowCol,
                subprogramCol: _subProgram,
                enableWdCol: _enableWdCol,
                bintableEnableWdCol: _bintableEnableWdCol,
                siteFlagPerSiteCol: _siteFlagPerSiteCol,
                failFlagCol: _failFlagCol,
                commentCol: _commentCol,
                moduleCol: _moduleCol,
                includeCol: _includeCol,
                groupCol: _groupCol,
                optionCol: _optionCol,
                jobs: _headerColumnDictionary.Keys.ToList(),
                enableModules: _enableModules);

            return mainFlowSheet;
        }

        private void ReadHeader()
        {
            string header;
            _startCol = _excelWorksheet.Dimension.Start.Column;
            _startRow = _excelWorksheet.Dimension.Start.Row;
            _endCol = _excelWorksheet.Dimension.End.Column;
            _endRow = _excelWorksheet.Dimension.End.Row;
            for (int i = _startCol; i <= _endCol; i++)
            {
                bool noHeader = false;
                header = GetCellValue(1, i);
                if (!MatchNameHeaderColumn(header, i) && !MatchEnableHeaderColumn(header, i))
                {
                    noHeader = true;
                }

                if (!noHeader)
                {
                    _columnIndex.Add(i);
                }
            }

            ReadJobColumns();

            BuildT0TxColumnDictionary();
        }

        private bool MatchNameHeaderColumn(string header, int i)
        {
            if (CellDiff.IsLiked(header, ConSource))
            {
                _sourceCol = i;
            }
            else if (CellDiff.IsLiked(header, ConSheetName1) || CellDiff.IsLiked(header, ConSheetName2))
            {
                _sheetNameCol = i;
            }
            else if (CellDiff.IsLiked(header, ConSubFlow))
            {
                _subFlowCol = i;
            }
            else if (CellDiff.IsLiked(header, ConSubProgram))
            {
                _subProgram = i;
            }
            else if (CellDiff.IsLiked(header, ConSiteFlagPerSite))
            {
                _siteFlagPerSiteCol = i;
            }
            else if (CellDiff.IsLiked(header, ConFailFlag))
            {
                _failFlagCol = i;
            }
            else if (CellDiff.IsLiked(header, ConModule))
            {
                _moduleCol = i;
            }
            else if (CellDiff.IsLiked(header, ConInclude))
            {
                _includeCol = i;
            }
            else if (CellDiff.IsLiked(header, ConGroup))
            {
                _groupCol = i;
            }
            else if (CellDiff.IsLiked(header, ConOption))
            {
                _optionCol = i;
            }
            else if (CellDiff.IsLiked(header, ConOptions))
            {
                _optionCol = i;
            }
            else
            {
                return false;
            }

            return true;
        }

        private bool MatchEnableHeaderColumn(string header, int i)
        {
            if (header.StartsWith(ConEnable1, StringComparison.CurrentCultureIgnoreCase) || Regex.IsMatch(header, ConEnable2, RegexOptions.IgnoreCase))
            {
                _enableWdCol = i;
            }
            else if (Regex.IsMatch(header, ConBintableEnable, RegexOptions.IgnoreCase))
            {
                _bintableEnableWdCol = i;
            }
            else
            {
                return false;
            }

            return true;
        }

        private void ReadJobColumns()
        {
            for (int i = _enableWdCol + 1; i <= _endCol; i++)
            {
                string header = GetCellValue(1, i);
                if (string.IsNullOrEmpty(header.Trim()))
                {
                    continue;
                }
                if (Regex.IsMatch(header, ConComment, RegexOptions.IgnoreCase))
                {
                    _commentCol = i;
                    _columnIndex.Add(i);
                    continue;
                }
                if (Regex.IsMatch(header, ConInclude, RegexOptions.IgnoreCase))
                {
                    _includeCol = i;
                    _columnIndex.Add(i);
                    continue;
                }
                if (_columnIndex.Exists(x => x.Equals(i)))
                {
                    continue;
                }

                foreach (string item in header.Split(','))
                {
                    string job = item.Split('(')[0].Trim();
                    _headerColumnDictionary.Add(job, i);
                }
            }
        }

        private void BuildT0TxColumnDictionary()
        {
            foreach (string headerCol in _headerColumnDictionary.Keys)
            {
                if (headerCol.Equals("CP1", StringComparison.CurrentCultureIgnoreCase) || headerCol.Equals("FT1", StringComparison.CurrentCultureIgnoreCase) || headerCol.Equals("WLFT1", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (_t0txColumnDictionary.ContainsKey("T0TX_Room"))
                    {
                        _t0txColumnDictionary["T0TX_Room"].Add(_headerColumnDictionary[headerCol]);
                    }
                    else
                    {
                        _t0txColumnDictionary.Add("T0TX_Room", new List<int> { _headerColumnDictionary[headerCol] });
                    }
                    continue;
                }
                if (headerCol.Equals("CP2", StringComparison.CurrentCultureIgnoreCase) || headerCol.Equals("FT2", StringComparison.CurrentCultureIgnoreCase) || headerCol.Equals("WLFT2", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (_t0txColumnDictionary.ContainsKey("T0TX_Hot"))
                    {
                        _t0txColumnDictionary["T0TX_Hot"].Add(_headerColumnDictionary[headerCol]);
                    }
                    else
                    {
                        _t0txColumnDictionary.Add("T0TX_Hot", new List<int> { _headerColumnDictionary[headerCol] });
                    }
                }
            }
        }

        private void ReadEnableModule()
        {
            for (int i = _startRow + 1; i <= _endRow; i++)
            {
                string module = _moduleCol == -1 ? "" : GetCellValue(i, _moduleCol).Trim();
                string include = GetCellValue(i, _includeCol).Trim();
                if (include.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                {
                    if (_enableModules == null)
                    {
                        _enableModules = new List<string>();
                    }

                    if (!_enableModules.Exists(x => x.Equals(module, StringComparison.OrdinalIgnoreCase)))
                    {
                        _enableModules.Add(module);
                    }
                }
            }
        }

        private void ReadDataMain()
        {
            foreach (KeyValuePair<string, int> keyValuePair in _headerColumnDictionary)
            {
                string header = keyValuePair.Key;
                int column = keyValuePair.Value;
                MainFlowBase mainFlowBase = new MainFlowBase();
                string[] jobItems = header.Split(':');
                string jobName = jobItems[0].ToUpper();
                string part = jobItems.Length > 1 ? jobItems[1] : "";
                mainFlowBase.JobName = jobName;
                MainFlowBase existJob = _mainFlowBases.Find(x => x.JobName.Equals(jobName));
                for (int i = _startRow + 1; i <= _endRow; i++)
                {
                    string oriSheetName = GetCellValue(i, _sheetNameCol).Trim();
                    string comment = "";
                    string module = _moduleCol == -1 ? "" : GetCellValue(i, _moduleCol).Trim();
                    if (existJob != null)
                    {
                        FlowSequenceNew targetSequence = existJob.SequencesNew.Find(x => x.OriSheetName.Equals(oriSheetName, StringComparison.OrdinalIgnoreCase));
                        if (!string.IsNullOrEmpty(part) && !existJob.AvalibleParts.Exists(x => x.Equals(part, StringComparison.OrdinalIgnoreCase)))
                        {
                            existJob.AvalibleParts.Add(part);
                        }

                        if (!string.IsNullOrEmpty(GetCellValue(i, column).Trim()))
                        {
                            string mergePart = string.Join(",", new List<string> { targetSequence.Part, part });
                            targetSequence.Part = mergePart;
                            targetSequence.Enable = true;
                        }
                    }
                    else
                    {
                        FlowSequenceNew sequence = new FlowSequenceNew
                        {
                            Enable = !string.IsNullOrEmpty(GetCellValue(i, column).Trim()),
                            Module = module,
                            Source = GetCellValue(i, _sourceCol).Trim(),
                            EnableWord = GetCellValue(i, _enableWdCol).Trim()
                        };
                        if (_groupCol != -1)
                        {
                            sequence.Group = GetCellValue(i, _groupCol).Trim();
                        }

                        if (_optionCol != -1)
                        {
                            sequence.Option = GetCellValue(i, _optionCol).Trim();
                        }
                        if (_bintableEnableWdCol != -1)
                        {
                            sequence.BintableEnableWord = GetCellValue(i, _bintableEnableWdCol).Trim();
                        }
                        if (_siteFlagPerSiteCol != -1)
                        {
                            sequence.SiteFlagPerSite = GetCellValue(i, _siteFlagPerSiteCol).Trim();
                        }

                        if (_failFlagCol != -1)
                        {
                            sequence.FailFlag = GetCellValue(i, _failFlagCol).Trim();
                        }

                        foreach (KeyValuePair<string, int> item in _headerColumnDictionary)
                        {
                            string value = GetCellValue(i, item.Value).Trim();
                            sequence.JobDic[item.Key] = value;
                        }
                        sequence.OriSheetName = oriSheetName;
                        string[] sheetItems = oriSheetName.Split(':');
                        sequence.SheetName = sheetItems[0];
                        if (sheetItems.Length > 1)
                        {
                            sequence.SubFlowName = sheetItems[1];
                        }

                        if (_subFlowCol != -1)
                        {
                            sequence.SubFlowName = GetCellValue(i, _subFlowCol).Trim();
                        }

                        if (_subProgram != -1)
                        {
                            sequence.SubProgramFlowName = GetCellValue(i, _subProgram).Trim();
                        }
                        sequence.Job = jobName;
                        if (sequence.Enable)
                        {
                            sequence.Part = part;
                        }

                        sequence.RowNum = i;
                        sequence.Comment = comment;
                        if (!string.IsNullOrEmpty(sequence.SheetName))
                        {
                            mainFlowBase.SequencesNew.Add(sequence);
                        }
                    }
                }
                if (existJob == null)
                {
                    if (!string.IsNullOrEmpty(part))
                    {
                        mainFlowBase.AvalibleParts.Add(part);
                    }
                    mainFlowBase.MainFlowName = "Main_Flow_" + mainFlowBase.JobName;
                    mainFlowBase.SheetName = _excelWorksheet.Name;
                    mainFlowBase.JobColIndex = column;
                    _mainFlowBases.Add(mainFlowBase);
                }
            }
        }

        private void ReadDataSub()
        {
            MainFlowBase mainFlowBase = new MainFlowBase { JobName = "All" };
            for (int i = _startRow + 1; i <= _endRow; i++)
            {
                FlowSequenceNew sequence = new FlowSequenceNew();
                if (_sourceCol != -1)
                {
                    sequence.Source = GetCellValue(i, _sourceCol).Trim();
                }

                if (_moduleCol != -1)
                {
                    sequence.Module = GetCellValue(i, _moduleCol).Trim();
                }

                if (_sheetNameCol != -1)
                {
                    string oriSheetName = GetCellValue(i, _sheetNameCol).Trim();
                    sequence.OriSheetName = oriSheetName;
                    string[] sheetItems = oriSheetName.Split(':');
                    sequence.SheetName = sheetItems[0];
                    if (sheetItems.Length > 1)
                    {
                        sequence.SubFlowName = sheetItems[1];
                    }
                }
                if (_subFlowCol != -1)
                {
                    sequence.SubFlowName = GetCellValue(i, _subFlowCol).Trim();
                }

                if (_subProgram != -1)
                {
                    sequence.SubProgramFlowName = GetCellValue(i, _subProgram).Trim();
                }

                if (_enableWdCol != -1)
                {
                    sequence.EnableWord = GetCellValue(i, _enableWdCol).Trim();
                }
                if (_bintableEnableWdCol != -1)
                {
                    sequence.BintableEnableWord = GetCellValue(i, _bintableEnableWdCol).Trim();
                }
                if (_siteFlagPerSiteCol != -1)
                {
                    sequence.SiteFlagPerSite = GetCellValue(i, _siteFlagPerSiteCol).Trim();
                }

                if (_failFlagCol != -1)
                {
                    sequence.FailFlag = GetCellValue(i, _failFlagCol).Trim();
                }

                foreach (KeyValuePair<string, int> keyValuePair in _headerColumnDictionary)
                {
                    string value = GetCellValue(i, keyValuePair.Value).Trim();
                    sequence.JobDic[keyValuePair.Key] = value;
                }
                if (_commentCol != -1)
                {
                    sequence.Comment = GetCellValue(i, _commentCol).Trim();
                }
                if (_bintableEnableWdCol != -1)
                {
                    sequence.BintableEnableWord = GetCellValue(i, _bintableEnableWdCol).Trim();
                }
                if (_groupCol != -1)
                {
                    sequence.Group = GetCellValue(i, _groupCol).Trim();
                }

                if (_optionCol != -1)
                {
                    sequence.Option = GetCellValue(i, _optionCol).Trim();
                }

                sequence.Enable = true;
                sequence.RowNum = i;

                if (!string.IsNullOrEmpty(sequence.SheetName))
                {
                    mainFlowBase.SequencesNew.Add(sequence);
                }
            }

            mainFlowBase.MainFlowName = "Main_Flow_Sub";
            mainFlowBase.SheetName = _excelWorksheet.Name;
            _mainFlowBases.Add(mainFlowBase);
        }

        private void ReadT0TXDataMain()
        {
            foreach (KeyValuePair<string, List<int>> keyValuePair in _t0txColumnDictionary)
            {
                string header = keyValuePair.Key;
                MainFlowBase mainFlowBase = new MainFlowBase();
                string[] jobItems = header.Split(':');
                string jobName = jobItems[0].ToUpper();
                string part = jobItems.Length > 1 ? jobItems[1] : "";
                mainFlowBase.JobName = jobName.Equals("T0TX_Room", StringComparison.CurrentCultureIgnoreCase) ? "FT1" : "FT2";
                MainFlowBase existJob = _mainFlowBases.Find(x => x.JobName.Equals(jobName));
                for (int i = _startRow + 1; i <= _endRow; i++)
                {
                    string oriSheetName = GetCellValue(i, _sheetNameCol).Trim();
                    string comment = "";
                    string module = _moduleCol == -1 ? "" : GetCellValue(i, _moduleCol).Trim();
                    if (existJob != null)
                    {
                        FlowSequenceNew targetSequence = existJob.SequencesNew.Find(x => x.OriSheetName.Equals(oriSheetName, StringComparison.OrdinalIgnoreCase));
                        if (!string.IsNullOrEmpty(part) && !existJob.AvalibleParts.Exists(x => x.Equals(part, StringComparison.OrdinalIgnoreCase)))
                        {
                            existJob.AvalibleParts.Add(part);
                        }

                        if (keyValuePair.Value.Any(x => !string.IsNullOrEmpty(GetCellValue(i, x).Trim())))
                        {
                            string mergePart = string.Join(",", new List<string> { targetSequence.Part, part });
                            targetSequence.Part = mergePart;
                            targetSequence.Enable = true;
                        }
                    }
                    else
                    {
                        FlowSequenceNew sequence = new FlowSequenceNew
                        {
                            Enable = keyValuePair.Value.Any(x => !string.IsNullOrEmpty(GetCellValue(i, x).Trim())),
                            Module = module,
                            Source = GetCellValue(i, _sourceCol).Trim(),
                            EnableWord = GetCellValue(i, _enableWdCol).Trim()
                        };
                        if (_groupCol != -1)
                        {
                            sequence.Group = GetCellValue(i, _groupCol).Trim();
                        }

                        if (_optionCol != -1)
                        {
                            sequence.Option = GetCellValue(i, _optionCol).Trim();
                        }
                        if (_bintableEnableWdCol != -1)
                        {
                            sequence.BintableEnableWord = GetCellValue(i, _bintableEnableWdCol).Trim();
                        }
                        if (_siteFlagPerSiteCol != -1)
                        {
                            sequence.SiteFlagPerSite = GetCellValue(i, _siteFlagPerSiteCol).Trim();
                        }

                        if (_failFlagCol != -1)
                        {
                            sequence.FailFlag = GetCellValue(i, _failFlagCol).Trim();
                        }

                        foreach (KeyValuePair<string, int> item in _headerColumnDictionary)
                        {
                            string value = GetCellValue(i, item.Value).Trim();
                            sequence.JobDic[item.Key] = value;
                        }
                        sequence.OriSheetName = oriSheetName;
                        string[] sheetItems = oriSheetName.Split(':');
                        sequence.SheetName = sheetItems[0];
                        if (sheetItems.Length > 1)
                        {
                            sequence.SubFlowName = sheetItems[1];
                        }

                        if (_subFlowCol != -1)
                        {
                            sequence.SubFlowName = GetCellValue(i, _subFlowCol).Trim();
                        }

                        if (_subProgram != -1)
                        {
                            sequence.SubProgramFlowName = GetCellValue(i, _subProgram).Trim();
                        }
                        sequence.Job = jobName.Equals("T0TX_Room", StringComparison.CurrentCultureIgnoreCase) ? "FT1" : "FT2";
                        if (sequence.Enable)
                        {
                            sequence.Part = part;
                        }

                        sequence.RowNum = i;
                        sequence.Comment = comment;
                        if (!string.IsNullOrEmpty(sequence.SheetName))
                        {
                            mainFlowBase.SequencesNew.Add(sequence);
                        }
                    }
                }
                if (existJob == null)
                {
                    if (!string.IsNullOrEmpty(part))
                    {
                        mainFlowBase.AvalibleParts.Add(part);
                    }
                    mainFlowBase.MainFlowName = "Main_Flow_" + jobName;
                    mainFlowBase.SheetName = _excelWorksheet.Name;
                    mainFlowBase.JobColIndex = keyValuePair.Value.First();
                    _mainFlowBases.Add(mainFlowBase);
                }
            }
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = _excelWorksheet.Cells[rowNumber, columnNumber].Value;
            if (value != null)
            {
                return value.ToString();
            }
            return "";
        }
    }
}
