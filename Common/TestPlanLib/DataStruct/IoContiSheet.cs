using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace TestPlanLib.DataStruct
{
    public class IoContiSheet : MySheet
    {
        public string Name = "";

        public List<IoContiRow> Rows { set; get; } = [];

        public void AddRow(IoContiRow ioContiRow)
        {
            Rows.Add(ioContiRow);
        }

        public bool TryGetFsDd(string pPinName, out string pFsDd)
        {
            pFsDd = "";
            bool exist = false;
            foreach (IoContiRow contiRow in Rows)
            {
                if (contiRow.BumpName.EqualsIgnoreCase(pPinName)
                    || contiRow.BallName.EqualsIgnoreCase(pPinName))
                {
                    pFsDd = contiRow.FsDd;
                    exist = true;
                }
            }

            return exist;
        }

        public string GetVoltage(string pPinName)
        {
            string voltage = "";

            foreach (IoContiRow contiRow in Rows)
            {
                if (contiRow.BumpName.EqualsIgnoreCase(pPinName)
                    || contiRow.BallName.EqualsIgnoreCase(pPinName))
                {
                    voltage = contiRow.IoVoltage;
                    if (!string.IsNullOrEmpty(contiRow.Chiplet))
                    {
                        voltage = voltage + "_" + contiRow.Chiplet;
                    }
                }
            }
            return voltage;
        }

        public List<string> GetPinList()
        {
            List<string> pinList = [];
            foreach (IoContiRow contiRow in Rows)
            {
                if (!pinList.Exists(p => p.EqualsIgnoreCase(contiRow.BumpName)))
                {
                    pinList.Add(contiRow.BumpName);
                }

                if (!pinList.Exists(p => p.EqualsIgnoreCase(contiRow.BallName)))
                {
                    pinList.Add(contiRow.BallName);
                }
            }
            return pinList;
        }

        public bool IsPinExist(string pPinName)
        {
            bool exist = false;
            foreach (IoContiRow contiRow in Rows)
            {
                if (contiRow.BumpName.EqualsIgnoreCase(pPinName)
                    || contiRow.BallName.EqualsIgnoreCase(pPinName))
                {
                    exist = true;
                }
            }

            return exist;
        }

        public static IoContiSheet? GenIoContiSheet(IoContiSheet ioContiSheet, ExcelWorksheet excelWorksheet)
        {
            if (ioContiSheet != null)
            {
                return ioContiSheet;
            }
            if (excelWorksheet != null)
            {
                IoLevelsSheetReader ioLevelsSheetReader = new IoLevelsSheetReader();
                IoLevelsSheet ioLevels = ioLevelsSheetReader.ReadSheet(excelWorksheet)!;
                return ioLevels.ConvertIoContiSheet();
            }
            return null;
        }

        public void CheckIoPins(PinMapSheet pinMapSheet, bool isIoLevel = false)
        {
            List<Pin> pins = pinMapSheet.GetIoContinuityPins();
            string sheetName = isIoLevel ? NeededSheets.IoLevels : NeededSheets.ContiIo;
            foreach (Pin pin in pins)
            {
                if (!GetPinList().Exists(x => x.EqualsIgnoreCase(pin.PinName)))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_06, sheetName, 0, 0, $"The pin {pin.PinName} can not be found in sheet {sheetName}", [pin.PinName, sheetName]);
                }
            }
        }

        public static void CheckIoPins(IoContiSheet ioContiSheet, Dictionary<string, string> ioIgnoreListDic, PinMapSheet pinMapSheet, bool isIoLevel = false)
        {
            string ioContiSheetName = isIoLevel ? NeededSheets.IoLevels : NeededSheets.ContiIo;
            if (ioIgnoreListDic == null)
            {
                ErrorReportManager.AddError(PreActionErrorType.E_MissingDocument_02, pinMapSheet.Name, 0, 0);
            }
            else
            {
                foreach (Pin pin in pinMapSheet.GetIoPins())
                {
                    if (!ioContiSheet.IsPinExist(pin.PinName))
                    {
                        foreach (KeyValuePair<string, string> regex in ioIgnoreListDic)
                        {
                            if (Regex.IsMatch(pin.PinName, regex.Key, RegexOptions.IgnoreCase))
                            {
                                string errorMessage = $"The I/O pin {pin.PinName} in {pinMapSheet.Name} can not be found in {ioContiSheetName} ({regex.Key}, {regex.Value})";

                                if (regex.Value.EqualsIgnoreCase("BYPASS"))
                                {
                                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_07, pinMapSheet.Name, 0, 0, $"The I/O pin {pin.PinName} in {pinMapSheet.Name} can not be found in {ioContiSheetName} ({regex.Key}, {regex.Value})", [pin.PinName, pinMapSheet.Name, ioContiSheetName, regex.Key, regex.Value]);
                                }
                                else if (regex.Value.EqualsIgnoreCase("WARNING"))
                                {
                                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_08, pinMapSheet.Name, 0, 0, $"The I/O pin {pin.PinName} in {pinMapSheet.Name} can not be found in {ioContiSheetName} ({regex.Key}, {regex.Value})", [pin.PinName, pinMapSheet.Name, ioContiSheetName, regex.Key, regex.Value]);
                                }
                                else
                                {
                                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_09, pinMapSheet.Name, 0, 0, $"The I/O pin {pin.PinName} in {pinMapSheet.Name} can not be found in {ioContiSheetName} ({regex.Key}, {regex.Value})", [pin.PinName, pinMapSheet.Name, ioContiSheetName, regex.Key, regex.Value]);
                                }

                                break;
                            }
                        }
                    }
                }
            }

            foreach (string pin in ioContiSheet.GetPinList())
            {
                if (!pinMapSheet.IsPinExist(pin))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_10, ioContiSheetName, 0, 0, $"The pin {pin} in {ioContiSheetName} can not be found in {pinMapSheet.Name}", [pin, ioContiSheetName, pinMapSheet.Name]);
                }
            }

            foreach (Pin pin in pinMapSheet.GetIoContinuityPins())
            {
                if (!ioContiSheet.GetPinList().Exists(x => x.EqualsIgnoreCase(pin.PinName)))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_11, ioContiSheetName, 0, 0, $"The pin {pin.PinName} can not be found in sheet {ioContiSheetName}", [pin.PinName, ioContiSheetName]);
                }
            }
        }
    }
}
