using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader
{
    public class SelSrmReader : TextInputReader
    {

        public List<SelSrmItem> SrmItems = new List<SelSrmItem>();
        private List<string> _patUseList = new List<string>();
        public static List<SelSrmItem> SelMappingList = new List<SelSrmItem>();
        private readonly List<SelSrmRow> _selSrmRows = new List<SelSrmRow>();
        private readonly Dictionary<string, int> _headerOrderDic = new Dictionary<string, int>();

        public SelSrmReader(string path) : base(path)
        {
            //using (var ep = new ExcelPackage(new FileInfo(path)))
            //{
            //    var wb = ep.Workbook;
            //    if (wb.Worksheets["SELSRM_Mapping_Table"] != null)
            //    {
            //        ReadMappingTable(wb.Worksheets["SELSRM_Mapping_Table"]);
            //        SelMappingList = SrmItems;
            //    }
            //    else
            //    {
            if (File.Exists(path))
            {
                FilePath = path;
            }
            //    }

            //}
            //_selSrmReader.ReadMappingTable(sh);
            //SelMappingList = _selSrmReader.SrmItems;
        }

        protected override void _Read(StreamReader sr)
        {
            SrmItems.Clear();
            _selSrmRows.Clear();
            _patUseList.Clear();

            string line = "";
            int i = 0;
            var header = new List<string>();
            SelSrmItem selsrmItem = new SelSrmItem();
            try
            {
                while ((line = sr.ReadLine()) != null && !line.ToLower().Contains("end"))
                {
                    //if (line.Contains("VDD_"))
                    {
                        string[] values = line.Split(new[] { '\t' });
                        if (i == 0)
                        {
                            header.AddRange(line.ToUpper().Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries));
                            i++;
                            continue;
                        }

                        if (!string.IsNullOrEmpty(values[header.IndexOf("PATTERN")]))
                        {
                            if (!_IsSameOrNull(values[header.IndexOf("STAGE")], selsrmItem.Stage) ||
                                !_IsSameOrNull(values[header.IndexOf("BLOCK")], selsrmItem.Block) ||
                                !_IsSameOrNull(values[header.IndexOf("PATTERN")], selsrmItem.Pattern))
                            {
                                selsrmItem = new SelSrmItem
                                {
                                    Stage = values[header.IndexOf("STAGE")],
                                    Block = values[header.IndexOf("BLOCK")],
                                    Pattern = values[header.IndexOf("PATTERN")]
                                };
                                SrmItems.Add(selsrmItem);
                            }
                        }

                        var selSrmRow = new SelSrmRow
                        {
                            RowNum = i, Bits = values[header.IndexOf("BITS")], LogicPins = values[header.IndexOf("LOGIC PINS")],
                            SramPins = values[header.IndexOf("SRAM PINS")],
                            Selsrm1 = values[header.IndexOf("SELSRM1")],
                            Selsrm0 = values[header.IndexOf("SELSRM0")]
                        };

                        selsrmItem.Rows.Add(selSrmRow);
                        i++;
                    }
                }
            }
            catch (Exception)
            {
            }

            _patUseList = SrmItems.Select(p => p.Pattern).ToList();
            SelMappingList = SrmItems;


        }

        private bool _IsSameOrNull(string input, string target)
        {
            if (string.IsNullOrEmpty(input))
            {
                return true;
            }
            else
            {
                return string.Equals(input, target, StringComparison.OrdinalIgnoreCase);
            }
        }

        public void ReadMappingTable(ExcelWorksheet sheet)
        {
            SrmItems.Clear();
            _selSrmRows.Clear();
            _patUseList.Clear();

            // read header
            int startRow = _SearchTableStartRow(sheet);
            int endRow = sheet.Dimension.End.Row;

            // early return
            if (startRow == endRow)
            {
                return;
            }

            _UpdateTableHeader(sheet, startRow);

            // read selsrm rows

            SelSrmItem selsrmItem = new SelSrmItem();

            for (int j = startRow + 1; j <= endRow; j++)
            {
                if (!string.IsNullOrEmpty(_GetCell(sheet, j, "pattern", _headerOrderDic)))
                {
                    selsrmItem = new SelSrmItem
                    {
                        Stage = _GetCell(sheet, j, "stage", _headerOrderDic),
                        Block = _GetCell(sheet, j, "block", _headerOrderDic),
                        Pattern = _GetCell(sheet, j, "pattern", _headerOrderDic)
                    };
                    SrmItems.Add(selsrmItem);
                }
                var selSrmRow = new SelSrmRow
                {
                    RowNum = j,
                    Bits = _GetCell(sheet, j, "bits", _headerOrderDic),
                    LogicPins = _GetCell(sheet, j, "logicpins", _headerOrderDic),
                    SramPins = _GetCell(sheet, j, "srampins", _headerOrderDic),
                    Selsrm0 = _GetCell(sheet, j, "selsrm0", _headerOrderDic),
                    Selsrm1 = _GetCell(sheet, j, "selsrm1", _headerOrderDic),
                };
                selsrmItem.Rows.Add(selSrmRow);
                _selSrmRows.Add(selSrmRow);
            }

            _patUseList = SrmItems.Select(p => p.Pattern).ToList();
        }

        public string GetPatternKey(string pattern)
        {
            return _patUseList.Exists(p => p.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                ? _patUseList.Find(p => p.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                : "";
        }

        private void _UpdateTableHeader(ExcelWorksheet sheet, int startrow)
        {
            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                string header = "";
                if (sheet.Cells[startrow, i].Value != null)
                {
                    header = sheet.Cells[startrow, i].Value.ToString().Trim();
                }

                if (header.Trim() != "")
                {
                    _headerOrderDic[header.Replace(" ", "").ToLower()] = i;
                }
            }
        }

        private static int _SearchTableStartRow(ExcelWorksheet sheet)
        {
            /* Return the row number of Column A == "Block" */
            for (int i = 1; i <= sheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= sheet.Dimension.End.Column; j++)
                {
                    if (sheet.Cells[i, j].Value != null && sheet.Cells[i, j].Value.ToString().Trim() == "Pattern")
                    {
                        return i;
                    }
                }
            }
            return sheet.Dimension.End.Row;
        }

        private static string _GetCell(ExcelWorksheet sheet, int j, string fieldName, Dictionary<string, int> headerOrder)
        {
            return headerOrder.Keys.Contains(fieldName)
                ? ExcelOperation.GetCellValue(sheet, j, headerOrder[fieldName])
                : "";
        }

    }
}
