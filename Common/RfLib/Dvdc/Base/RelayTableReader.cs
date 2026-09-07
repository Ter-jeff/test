using System.Collections.Generic;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace RfLib.Dvdc.Base
{
    public partial class RelayTableReader
    {
        [GeneratedRegex("Path Name", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        private int _colPathNameIndex = -1;
        private readonly Dictionary<int, Component> _indexComponents = [];
        public List<RelayPathItem> RelayPathItems = [];
        public void Read(ExcelWorksheet excelWorksheet)
        {
            int rowindex = 1;
            ReadHeader(excelWorksheet, ref rowindex);
            ReadContent(excelWorksheet, rowindex);
        }

        //Path Name
        private void ReadHeader(ExcelWorksheet excelWorksheet, ref int rowIndex)
        {
            //Path names	LUT
            #region get path index
            for (rowIndex = 1; rowIndex <= excelWorksheet.Dimension.Rows; rowIndex++)
            {
                if (_colPathNameIndex != -1)
                {
                    break;
                }

                for (int j = 1; j <= excelWorksheet.Dimension.Columns; j++)
                {
                    if (string.IsNullOrEmpty(excelWorksheet.Cells[rowIndex, j].Text))
                    {
                        continue;
                    }

                    if (MyRegex().IsMatch(excelWorksheet.Cells[rowIndex, j].Text))
                    {
                        _colPathNameIndex = j;
                    }
                }
            }
            #endregion

            #region get components index
            for (int j = _colPathNameIndex + 1; j <= excelWorksheet.Dimension.Columns; j++)
            {
                if (string.IsNullOrEmpty(excelWorksheet.Cells[rowIndex, j].Text) || string.IsNullOrEmpty(excelWorksheet.Cells[rowIndex - 1, j].Text))
                {
                    continue;
                }

                var component = new Component(excelWorksheet.Cells[rowIndex - 1, j].Text, excelWorksheet.Cells[rowIndex, j].Text);

                _indexComponents.Add(j, component);
            }
            #endregion

            rowIndex++;
        }

        private void ReadContent(ExcelWorksheet excelWorksheet, int rowIndex)
        {
            var recordPaths = new List<string>();
            for (int i = rowIndex; i <= excelWorksheet.Dimension.Rows; i++)
            {
                //var path = pathRename(sheet.Cells[i, _colPathNameIndex].Text);
                string path = excelWorksheet.Cells[i, _colPathNameIndex].Text;
                if (!string.IsNullOrEmpty(path))
                {
                    var item = new RelayPathItem { PathName = path };
                    if (recordPaths.Contains(item.PathName))
                    {
                        continue;
                    }

                    foreach (KeyValuePair<int, Component> indexComponent in _indexComponents)
                    {
                        if (!string.IsNullOrEmpty(excelWorksheet.Cells[i, indexComponent.Key].Text))
                        {
                            Component itemComponent = indexComponent.Value.Clone();
                            itemComponent.Status = excelWorksheet.Cells[i, indexComponent.Key].Text;
                            item.ComponentInfos.Add(itemComponent);
                        }
                    }
                    recordPaths.Add(item.PathName);
                    RelayPathItems.Add(item);
                    if (item.ComponentInfos.Count == 0)
                    {
                    }
                }
            }
        }
    }
}
