using System;
using System.Collections.Generic;
using System.IO;

using Automation.InputManager.Data;
using Automation.PreCheck.PreChecks.PreAction;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility.Tester;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.Static;

namespace Automation.InputManager
{
    public class PreActionInputManager : InputManagerBase<PreActionInputData>
    {
        private readonly string _regPinMap = NeededSheets.PinMap;
        private readonly string _regPinGroup = NeededSheets.IoGroup;

        public PreActionInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override PreActionInputData Read()
        {
            var result = new PreActionInputData();
            string deviceType = LocalSpecs.Options.Device.ToString();
            string file = Path.Combine(AppContext.BaseDirectory, "Config", "Tester", "TesterConfig_" + deviceType + ".xml");
            if (!File.Exists(file))
            {
                file = Path.Combine(AppContext.BaseDirectory, "Config", "Tester", "TesterConfig_Default.xml");
            }

            Dictionary<string, TesterConfig> testerConfigs = TesterConfigReader.GetTesterConfigs(file);

            foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                string sheetName = sheet.Name;
                if (sheetName.Equals(_regPinMap, StringComparison.OrdinalIgnoreCase))
                {
                    PinMapSheet ioPinMap = new ReadPinMapSheet().ReadSheet(sheet);
                    ioPinMap.Name = "PinMap";
                    foreach (Pin pin in ioPinMap.PinList)
                    {
                        pin.ColumnA = $"{pin.SheetName} : {pin.RowNum}";
                    }
                    foreach (PinGroup group in ioPinMap.GroupList)
                    {
                        foreach (Pin pin in group.PinList)
                        {
                            pin.ColumnA = $"{pin.SheetName} : {pin.RowNum}";
                        }
                    }
                    result.PinMapSheet = ioPinMap;
                }
                else if (sheetName.Equals(_regPinGroup, StringComparison.OrdinalIgnoreCase))
                {
                    result.PinGroupSheet = sheet;
                }
                else if (sheetName.ContainsIgnoreCase("ChannelMap"))
                {
                    ChannelMapSheet channelMapSheet = ReadChanMapSheet.GetSheet(sheet, testerConfigs);
                    result.ChannelMapSheets.Add(channelMapSheet);
                }
                else if (sheetName.StartsWith(NeededSheets.ContiIo, StringComparison.CurrentCultureIgnoreCase))
                {
                    var preActionPreCheckIoContinuity = new PreActionPreCheckIoContinuity(EpWorkbook.TestPlanWorkbook, sheetName);
                    if (preActionPreCheckIoContinuity.CheckMain())
                    {
                        var ioContiReader = new IoContiReader();
                        result.IoContinuity = ioContiReader.ReadSheet(sheet);
                    }
                }
                else if (NeededSheets.IsTestSettingSheetName(sheetName, LocalSpecs.CurrentProject))
                {
                    var preActionPreCheckTestSetting = new PreActionPreCheckTestSetting(EpWorkbook.TestPlanWorkbook, sheetName);
                    if (!preActionPreCheckTestSetting.CheckMain())
                    {
                        preActionPreCheckTestSetting.HasSpecialUnit();
                    }
                }
            }

            return result;
        }
    }
}
