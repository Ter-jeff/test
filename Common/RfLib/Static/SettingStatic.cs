using System.IO;

using Automation.Const;
using Automation.Static;

using CommonLib.Extension;

using OfficeOpenXml;

using RfLib.InstrumentSetup;

namespace RfLib.Static
{
    internal class SettingStatic
    {
        private static InstrumentConfigSheet? _instrumentSheet;
        public static InstrumentConfigSheet? InstrumentSheet
        {
            get
            {
                if (_instrumentSheet == null && Automation.Static.SettingStatic.BasicConfigWorkbook != null)
                {
                    ExcelWorksheet instrumentSetupSheet = Automation.Static.SettingStatic.BasicConfigWorkbook.Worksheets[SheetConst.InstrumentSetup];
                    if (instrumentSetupSheet != null)
                    {
                        instrumentSetupSheet.ExportToTxt(Path.Combine(FolderStructure.DirHardIp, instrumentSetupSheet.Name + ".txt"));
                        TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirHardIp, instrumentSetupSheet.Name);
                        _instrumentSheet = new InstrumentConfigReader().ReadSheet(instrumentSetupSheet);
                        return _instrumentSheet;
                    }
                }
                return _instrumentSheet;
            }
        }

        public static void Clear()
        {
            _instrumentSheet = null;
        }
    }
}
