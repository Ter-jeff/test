using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public class ReadPinMapSheet : IgxlSheetReader<PinMapSheet>
    {
        private const int StartRowIndex = 3;

        public override EnumSheetType SupportedType => EnumSheetType.DTPinMap;

        public override PinMapSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            try
            {
                if (excelWorksheet.Dimension == null)
                {
                    return null!;
                }

                int endRow = excelWorksheet.Dimension.End.Row;
                var pinMapSheet = new PinMapSheet(excelWorksheet);
                var groupList = new Dictionary<string, PinGroup>();
                for (int i = 4; i <= endRow; i++)
                {
                    string columnA = EpplusExtensions.GetCellValue(excelWorksheet, i, 1);
                    string groupName = EpplusExtensions.GetCellValue(excelWorksheet, i, 2);
                    string pinName = EpplusExtensions.GetCellValue(excelWorksheet, i, 3).ToUpper();
                    string pinType = EpplusExtensions.GetCellValue(excelWorksheet, i, 4);
                    string comment = EpplusExtensions.GetCellValue(excelWorksheet, i, 5);
                    if (string.IsNullOrEmpty(groupName) && !string.IsNullOrEmpty(pinName))
                    {
                        var pin = new Pin(pinName, pinType)
                        {
                            SheetName = excelWorksheet.Name,
                            RowNum = i,
                            ColumnA = columnA
                        };
                        pinMapSheet.AddRow(pin);
                    }
                    else if (!string.IsNullOrEmpty(groupName))
                    {
                        if (!groupList.TryGetValue(groupName, out PinGroup? value))
                        {
                            var pinGroup = new PinGroup(groupName, pinType)
                            {
                                RowNum = i,
                                ColumnA = columnA,
                            };
                            var pin = new Pin(pinName, pinType)
                            {
                                SheetName = excelWorksheet.Name,
                                RowNum = i,
                                ColumnA = columnA,
                                Comment = comment
                            };
                            pinGroup.AddPin(pin);
                            pinMapSheet.AddRow(pinGroup);
                            groupList.Add(groupName, pinGroup);
                        }
                        else
                        {
                            var pin = new Pin(pinName, pinType)
                            {
                                SheetName = excelWorksheet.Name,
                                RowNum = i,
                                ColumnA = columnA,
                                Comment = comment
                            };
                            value.AddPin(pin);
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                return pinMapSheet;

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to read pin map sheet '{excelWorksheet.Name}'.", ex);
            }
        }

        public override PinMapSheet ReadSheet(Stream stream, string sheetName)
        {
            var pinMapSheet = new PinMapSheet(sheetName);
            bool isBackup = false;
            int i = 1;
            var groups = new Dictionary<string, PinGroup>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    string? line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        if (line != null)
                        {
                            string[] arr = line.Split('\t');
                            string columnA = arr[0];
                            string groupName = arr[1];
                            string pinName = arr[2];
                            string pinType = arr[3];
                            string comment = arr[4];
                            if (string.IsNullOrEmpty(pinName))
                            {
                                isBackup = true;
                                continue;
                            }

                            var pin = new Pin(pinName, pinType)
                            {
                                RowNum = i,
                                SheetName = sheetName,
                                IsBackup = isBackup,
                                ColumnA = columnA,
                                Comment = comment
                            };
                            if (string.IsNullOrEmpty(groupName) && !string.IsNullOrEmpty(pinName))
                            {
                                pinMapSheet.AddRow(pin);
                            }
                            else if (!string.IsNullOrEmpty(groupName))
                            {
                                if (!groups.TryGetValue(groupName, out PinGroup? value))
                                {
                                    var pinGroup = new PinGroup(groupName)
                                    {
                                        RowNum = i,
                                        ColumnA = columnA,
                                    };
                                    var pins = new List<Pin> { pin };
                                    pinGroup.AddPins(pins);
                                    pinMapSheet.AddRow(pinGroup);
                                    groups.Add(groupName, pinGroup);
                                }
                                else
                                {
                                    value.AddPin(pin);
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                    i++;
                }
            }

            return pinMapSheet;
        }
    }
}
