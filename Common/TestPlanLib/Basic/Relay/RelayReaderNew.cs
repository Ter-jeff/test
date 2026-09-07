using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.Basic.Relay
{
    public partial class RelayReaderNew
    {
        private const string ConHeaderRelayItem = "Relay name";

        [GeneratedRegex(@"Item\s*Name", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"Relay\s*name", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex1 = MyRegex1();
        private ExcelWorksheet? _excelWorksheet;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private readonly Dictionary<int, string> _header = [];

        public RelayTableNew? GetRelayTable(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            RelayTableNew relayTable = new RelayTableNew();

            _excelWorksheet = excelWorksheet;

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            // Set Content Value
            //set header List
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrValue = EpplusExtensions.GetCellValue(excelWorksheet, _startRowNumber, i);
                if (!string.IsNullOrEmpty(lStrValue))
                {
                    _header.Add(i, lStrValue);
                }
            }

            for (int i = _startRowNumber + 2; i <= _endRowNumber; i++)
            {
                var relayItem = new RelayItem();

                for (int j = _startColNumber; j <= _endColNumber; j++)
                {
                    string lStrValue = EpplusExtensions.GetCellValue(excelWorksheet, i, j);
                    if (string.IsNullOrEmpty(lStrValue))
                    {
                        continue;
                    }

                    if (!_header.TryGetValue(j, out string? relatedKey))
                    {
                        continue;
                    }

                    //Relay name	Item Name
                    if (_regex1.IsMatch(relatedKey))
                    {
                        relayItem.SubFlow = lStrValue;
                    }
                    else if (_regex.IsMatch(relatedKey))
                    {
                        relayItem.Item = lStrValue;
                    }
                    else
                    {
                        relayItem.RelayStatusDictionary.Add(relatedKey, lStrValue);
                    }
                }
                relayTable.RelayItems.Add(relayItem);
            }
            return relayTable;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderRelayItem))
                    {
                        _startRowNumber = i;
                        _startColNumber = j;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet!.Dimension != null)
            {
                _startColNumber = _excelWorksheet!.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet!.Dimension.Start.Row;
                _endColNumber = _excelWorksheet!.Dimension.End.Column;
                _endRowNumber = _excelWorksheet!.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _header.Clear();
        }
    }

    public class RelayTableNew
    {
        public List<string> Type = [];
        public List<RelayItem> RelayItems = [];

        public DataTable ConvertOldTable()
        {
            var relayFlows =
                RelayItems.Where(p => string.IsNullOrEmpty(p.Item) && !string.IsNullOrEmpty(p.SubFlow)).ToList();
            var utilities = relayFlows.SelectMany(p => p.RelayStatusDictionary).Select(p => p.Key).Distinct().ToList();
            var result = new DataTable();
            result.Columns.Add("Conti-Specify");
            result.Columns.Add("Relay Name");
            foreach (string utility in utilities)
            {
                result.Columns.Add(utility);
            }

            foreach (RelayItem relayInfo in relayFlows)
            {
                DataRow dr = result.NewRow();
                dr[1] = relayInfo.SubFlow;
                foreach (KeyValuePair<string, string> utility in relayInfo.RelayStatusDictionary)
                {
                    if (!utilities.Contains(utility.Key))
                    {
                        continue;
                    }

                    int columnIndex = utilities.FindIndex(p => p == utility.Key) + 2;
                    dr[columnIndex] = utility.Value;
                }

                result.Rows.Add(dr);
            }

            return result;
        }
    }

    public class RelayItem
    {
        public string SubFlow = "";
        public string Item = "";
        public Dictionary<string, string> RelayStatusDictionary = [];
    }
}
