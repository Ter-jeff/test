using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace TestPlanLib.Basic
{
    public class IoInfoReader
    {
        private const string ConPinGrpSymbol = "PinGrpName";
        private ExcelWorksheet? _currentSheet;
        private Dictionary<string, int> _titleDic = [];
        private readonly Dictionary<string, (int, int)> _headIndexDic = new(StringExtensions.IgnoreCase);
        private readonly Dictionary<string, List<IoInfoRow>> _ioInfoDic = new(StringExtensions.IgnoreCase);

        public IoInfoSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            _currentSheet = excelWorksheet;
            if (_currentSheet == null)
            {
                return null;
            }
            FindBlocks();

            foreach (KeyValuePair<string, (int, int)> head in _headIndexDic)
            {
                _ioInfoDic[head.Key] = ReadBlock(head);
            }

            if (_ioInfoDic.TryGetValue("Common", out List<IoInfoRow>? value))
            {
                foreach (string block in _ioInfoDic.Keys.Where(x => !x.StartsWithIgnoreCase("Levels_")))
                {
                    FillBlockData(_ioInfoDic[block], value);
                }
            }

            IoInfoSheet result = new IoInfoSheet(_currentSheet.Name, _ioInfoDic);

            return result;
        }

        private List<IoInfoRow> ReadBlock(KeyValuePair<string, (int, int)> blockIndex)
        {
            List<IoInfoRow> result = [];

            int bodyStartRow = blockIndex.Value.Item1;
            int bodyStartCol = blockIndex.Value.Item2;
            _titleDic = GetTittle(bodyStartRow, bodyStartCol);
            bool columnDefault = false;
            int nvIndex = FindTitleIndex(_titleDic, "NV");
            if (nvIndex == 0)
            {
                columnDefault = true;
            }

            int hvIndex = FindTitleIndex(_titleDic, "HV");
            if (hvIndex == 0)
            {
                columnDefault = true;
            }

            int lvIndex = FindTitleIndex(_titleDic, "LV");
            if (lvIndex == 0)
            {
                columnDefault = true;
            }

            int vilIndex = FindTitleIndex(_titleDic, "Vil");
            if (vilIndex == 0)
            {
                columnDefault = true;
            }

            int vilhIdex = FindTitleIndex(_titleDic, "Vih");
            if (vilhIdex == 0)
            {
                columnDefault = true;
            }

            int volIndex = FindTitleIndex(_titleDic, "Vol");
            if (volIndex == 0)
            {
                columnDefault = true;
            }

            int vohIndex = FindTitleIndex(_titleDic, "Voh");
            if (vohIndex == 0)
            {
                columnDefault = true;
            }

            int iolIndex = FindTitleIndex(_titleDic, "Iol");
            int iohIndex = FindTitleIndex(_titleDic, "Ioh");
            int vtIndex = FindTitleIndex(_titleDic, "Vt");
            int driverModeIndex = FindTitleIndex(_titleDic, "DriverMode");
            int timeDomainIndex = FindTitleIndex(_titleDic, "Time Domain");
            int vchIndex = FindTitleIndex(_titleDic, "Vch");
            int vclIndex = FindTitleIndex(_titleDic, "Vcl");
            if (columnDefault)
            {
                throw new Exception("IoInfo sheet default column.");
            }

            for (int row = bodyStartRow + 1; row <= _currentSheet!.Dimension.Rows; row++)
            {
                IoInfoRow info = new IoInfoRow
                {
                    PinGrpName = EpplusExtensions.GetCellValue(_currentSheet, row, bodyStartCol)
                };
                if (info.PinGrpName?.Length == 0)
                {
                    break;
                }
                info.Nv = EpplusExtensions.GetCellValue(_currentSheet, row, nvIndex);
                info.Hv = EpplusExtensions.GetCellValue(_currentSheet, row, hvIndex);
                info.Lv = EpplusExtensions.GetCellValue(_currentSheet, row, lvIndex);
                info.Vil = EpplusExtensions.GetCellValue(_currentSheet, row, vilIndex);
                info.Vih = EpplusExtensions.GetCellValue(_currentSheet, row, vilhIdex);
                info.Vol = EpplusExtensions.GetCellValue(_currentSheet, row, volIndex);
                info.Voh = EpplusExtensions.GetCellValue(_currentSheet, row, vohIndex);
                if (iolIndex != 0)
                {
                    info.Iol = EpplusExtensions.GetCellValue(_currentSheet, row, iolIndex);
                }

                if (iohIndex != 0)
                {
                    info.Ioh = EpplusExtensions.GetCellValue(_currentSheet, row, iohIndex);
                }
                if (vtIndex != 0)
                {
                    info.Vt = EpplusExtensions.GetCellValue(_currentSheet, row, vtIndex);
                }

                if (driverModeIndex != 0)
                {
                    info.DriverMode = EpplusExtensions.GetCellValue(_currentSheet, row, driverModeIndex);
                }

                if (timeDomainIndex != 0)
                {
                    info.TimeDomain = EpplusExtensions.GetCellValue(_currentSheet, row, timeDomainIndex);
                }

                if (vchIndex != 0)
                {
                    info.Vch = EpplusExtensions.GetCellValue(_currentSheet, row, vchIndex);
                }

                if (vclIndex != 0)
                {
                    info.Vcl = EpplusExtensions.GetCellValue(_currentSheet, row, vclIndex);
                }

                info.Type = blockIndex.Key;
                result.Add(info);
            }
            _titleDic.Clear();
            return result;
        }

        private static void FillBlockData(List<IoInfoRow> blockData, List<IoInfoRow> sourceData)
        {
            List<string> sourcePins = sourceData.ConvertAll(x => x.PinGrpName);
            foreach (IoInfoRow row in blockData)
            {
                if (!sourcePins.Exists(x => x.EqualsIgnoreCase(row.PinGrpName)))
                {
                    continue;
                }

                IoInfoRow sourceRow = sourceData.Find(x => x.PinGrpName.EqualsIgnoreCase(row.PinGrpName))!;
                PropertyInfo[] props = row.GetType().GetProperties();
                foreach (PropertyInfo prop in props)
                {
                    if (prop.Name == "NV")
                    {
                        string? sourceValue = sourceRow.GetType().GetProperty(prop.Name)!.GetValue(sourceRow, null)!.ToString();
                        if (string.IsNullOrEmpty(sourceValue))
                        {
                            throw new Exception("IoInfo sheet default NV value in Common block");
                        }

                        prop.SetValue(row, sourceValue, null);
                        continue;
                    }
                    if (!prop.CanWrite)
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty((string?)prop.GetValue(row, null)))
                    {
                        object? sourceValue = sourceRow.GetType().GetProperty(prop.Name)!.GetValue(sourceRow, null);
                        prop.SetValue(row, sourceValue, null);
                    }
                }
            }
        }

        private void FindBlocks()
        {
            for (int row = 2; row <= _currentSheet!.Dimension.Rows; row++)
            {
                for (int col = 1; col <= _currentSheet.Dimension.Columns; col++)
                {
                    string firstText = EpplusExtensions.GetCellValue(_currentSheet, row - 1, col);
                    string secondText = EpplusExtensions.GetCellValue(_currentSheet, row, col);

                    if (!string.IsNullOrEmpty(firstText)
                        && secondText.EqualsIgnoreCase(ConPinGrpSymbol))
                    {
                        if (!_headIndexDic.ContainsKey(firstText))
                        {
                            _headIndexDic[firstText] = (row, col);
                        }
                    }
                }
            }
        }

        private Dictionary<string, int> GetTittle(int bodyStartRow, int bodyStartCol)
        {
            Dictionary<string, int> result = [];
            for (int i = bodyStartCol + 1; i <= _currentSheet!.Dimension.Columns; i++)
            {
                if (EpplusExtensions.GetCellValue(_currentSheet, bodyStartRow, i)?.Length == 0)
                {
                    break;
                }

                result.Add(EpplusExtensions.GetCellValue(_currentSheet, bodyStartRow, i), i);
            }
            return result;
        }

        private static int FindTitleIndex(Dictionary<string, int> title, string titleName)
        {
            Regex titleReg = new Regex(titleName, RegexOptions.IgnoreCase);
            KeyValuePair<string, int> entry = title.FirstOrDefault(x => titleReg.IsMatch(x.Key));
            return entry.Key != null ? entry.Value : 0;
        }
    }
}
