using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.Relay
{
    public class RelayReader : MySheetReader<RelaySheet>
    {
        private const string ConHeaderSignals = "signals";
        private const string ConHeaderCircuit = "circuit";
        private const string ConHeaderSignalName = "signal name";
        private const string ConHeaderPn = "PN";
        private readonly Dictionary<int, string> _otherHeaders = new Dictionary<int, string>();

        private int _indexSignals = -1;
        private int _indexCircuit = -1;
        private int _indexSignalName = -1;
        private int _indexPn = -1;

        public override RelaySheet ReadSheet(ExcelWorksheet worksheet)
        {
            string sheetName = worksheet.Name;

            var relayNewSheet = new RelaySheet(sheetName);

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                relayNewSheet.AddDimensionError();
                return relayNewSheet;
            }

            if (!GetFirstHeaderPosition())
            {
                relayNewSheet.AddFirstHeaderError(ConHeaderPn);
                return relayNewSheet;
            }

            GetHeaderIndex();

            relayNewSheet = ReadSheet(sheetName);

            return relayNewSheet;
        }

        private RelaySheet ReadSheet(string sheetName)
        {
            var relayNewSheet = new RelaySheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new RelayNewRow { RowNum = i };
                if (_indexSignals != -1)
                {
                    row.Signals = ExcelWorksheet.GetCellValue(i, _indexSignals).Trim();
                }

                if (_indexCircuit != -1)
                {
                    row.Circuit = ExcelWorksheet.GetCellValue(i, _indexCircuit).Trim();
                }

                if (_indexSignalName != -1)
                {
                    row.SignalName = ExcelWorksheet.GetCellValue(i, _indexSignalName).Trim();
                }

                if (_indexPn != -1)
                {
                    row.Pn = ExcelWorksheet.GetCellValue(i, _indexPn).Trim();
                }

                foreach (int otherindex in _otherHeaders.Keys)
                {
                    row.OtherInfo.Add(otherindex, ExcelWorksheet.GetCellValue(i, otherindex).Trim().ToUpper());
                }
                relayNewSheet.Rows.Add(row);
            }

            relayNewSheet.IndexSignals = _indexSignals;
            relayNewSheet.IndexCircuit = _indexCircuit;
            relayNewSheet.IndexSignalName = _indexSignalName;
            relayNewSheet.IndexPn = _indexPn;
            string regRelay = "^Relay";
            var relayItems = new List<RelayItemNew>();
            foreach (KeyValuePair<int, string> relayFlow in _otherHeaders)
            {
                string relayName = relayFlow.Value;
                if (Regex.IsMatch(relayName, regRelay, RegexOptions.IgnoreCase))
                {
                    var infos = relayNewSheet.Rows.GroupBy(p => p.OtherInfo[relayFlow.Key]).ToDictionary(p => p.Key, p => p.Select(q => q.Pn).ToList());
                    if (infos.Keys.Any(x => x != "" && x != "ON" && x != "OFF"))
                    {
                        continue;
                    }

                    RelayItemNew target = relayItems.FirstOrDefault(p => p.Module.Equals(relayName, StringComparison.OrdinalIgnoreCase));
                    if (target == null)
                    {
                        target = new RelayItemNew(relayName);
                        relayItems.Add(target);
                    }

                    if (infos.TryGetValue("ON", out List<string> info))
                    {
                        target.On = info;
                    }

                    if (infos.TryGetValue("OFF", out List<string> info1))
                    {
                        target.Off = info1;
                    }
                }
            }
            relayNewSheet.RelayItems = relayItems;
            return relayNewSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.Equals(ConHeaderSignals, StringComparison.OrdinalIgnoreCase))
                {
                    _indexSignals = i;
                    continue;
                }
                if (header.Equals(ConHeaderCircuit, StringComparison.OrdinalIgnoreCase))
                {
                    _indexCircuit = i;
                    continue;
                }
                if (header.Equals(ConHeaderSignalName, StringComparison.OrdinalIgnoreCase))
                {
                    _indexSignalName = i;
                    continue;
                }
                if (header.Equals(ConHeaderPn, StringComparison.OrdinalIgnoreCase))
                {
                    _indexPn = i;
                }
                else
                {
                    _otherHeaders.Add(i, header);
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().Equals(ConHeaderPn, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class RelaySheet : MySheet
    {
        public List<RelayNewRow> Rows { set; get; }

        public int IndexSignals = -1;
        public int IndexCircuit = -1;
        public int IndexSignalName = -1;
        public int IndexPn = -1;
        public List<RelayItemNew> RelayItems;
        #region Constructor
        public RelaySheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<RelayNewRow>();
            RelayItems = new List<RelayItemNew>();
        }
        #endregion
    }

    public class RelayNewRow
    {
        public int RowNum { get; set; }
        public string Signals { set; get; }
        public string Circuit { set; get; }
        public string SignalName { set; get; }
        public string Pn { set; get; }
        public Dictionary<int, string> OtherInfo { get; set; } = new Dictionary<int, string>();
    }
    public class RelayItemNew
    {
        public string Module;
        public List<string> On;
        public List<string> Off;
        public RelayItemNew(string module)
        {
            Module = module;
            On = new List<string>();
            Off = new List<string>();
        }
    }
}
