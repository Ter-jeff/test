using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Utility.TpUpdate.HardIPEnableWordsUpdate;

using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Efuse.Input;
using TestPlanLib.Static;

namespace Automation.InputManager
{
    public class HardIpInputManager : InputManagerBase<HardIpInputData>
    {
        private readonly HardIpParaData _paraData;
        private readonly EnumBlock _block;
        public static readonly Regex RegexHardIp = new Regex("^(HARDIP_|PLLDEBUG_)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly Regex RegexIds = new Regex("^DCTEST_(?!Continuity).*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly Regex RegexRtos = new Regex("^RTOS_(?!Table|configuration).*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly Regex RegexSpirom = new Regex("_SPIROM", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public HardIpInputManager(ExcelWorkbook excelWorkbook, HardIpParaData hardIpParaData) : base(excelWorkbook)
        {
            _paraData = hardIpParaData;
            _block = hardIpParaData.Block;
        }

        public override HardIpInputData Read()
        {
            var result = new HardIpInputData(_paraData);
            Response.Report("Reading PinMap file ...", percentage: 10);
            IoPinMapReader.ReadPinMap();

            Response.Report("Reading BasicConfigure Setting file ...", percentage: 20);
            var blockMapReader = new BlockMappingReader(result);
            (result.VbtNameMapping, _) = blockMapReader.ConfigReader();

            //Move non Igxl sheet from TestPlan
            foreach (ExcelWorksheet ws in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                if (Regex.IsMatch(ws.Name, "^SweepListTable_|DSSCSetupSheet|CaptoEfuseSetup|CPP|RAK_|" + NeededSheets.HardIRegAssign, RegexOptions.IgnoreCase))
                {
                    if (TestProgram.NonIgxlSheetsList.SheetList.Exists(p => Path.GetFileNameWithoutExtension(p).Equals(ws.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    ws.ExportToTxt(Path.Combine(FolderStructure.DirHardIp, ws.Name + ".txt"));
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirHardIp, ws.Name);
                }
            }

            result.HardIpDcSheet = TestPlanStatic.HardIpDcSheet;
            result.IdsMappingSheet = TestPlanStatic.IdsMappingSheet;
            result.PowerInfoSheet = TestPlanStatic.PowerInfoSheet;
            result.PowerMergeSheet = TestPlanStatic.PowerMergeSheet;

            if (File.Exists(LocalSpecs.TtrSummaryFileName))
            {
                result.TtrTables = new TtrToolReader().ReadEnableWordTables(LocalSpecs.TtrSummaryFileName);
            }

            Response.Report("Reading Test Plan file ...", percentage: 80);
            IEnumerable<ExcelWorksheet> sheets = EpWorkbook.TestPlanWorkbook.Worksheets.Where(ws => !ws.Name.Equals("HardIp_Dc", StringComparison.OrdinalIgnoreCase) && !ws.Name.StartsWith(NeededSheets.ContiDcTest, StringComparison.OrdinalIgnoreCase));

            List<BitDefTable> efuseBitDefTables = TestPlanStatic.BitDefTables;
            if (efuseBitDefTables != null)
            {
                var efuseBitDefBw = new Dictionary<string, string>();
                foreach (BitDefTable efuseBitDefTable in efuseBitDefTables)
                {
                    foreach (BitDefRow efuseBitDefTableRow in efuseBitDefTable.Rows)
                    {
                        string bankEfuseBitDef = efuseBitDefTableRow.RowData[efuseBitDefTable.BankEfuseBitDefIdx];
                        string bitWidth = efuseBitDefTableRow.RowData[efuseBitDefTable.BitWidthIdx];
                        efuseBitDefBw[bankEfuseBitDef] = bitWidth;
                    }
                }
                result.EfuseBitDefBw = efuseBitDefBw;
            }

            switch (_block)
            {
                case EnumBlock.HardIp:
                    sheets = sheets.Where(ws => RegexHardIp.IsMatch(ws.Name));
                    break;
                case EnumBlock.Ids:
                    sheets = sheets.Where(ws => RegexIds.IsMatch(ws.Name));
                    break;
                case EnumBlock.Rtos:
                    sheets = sheets.Where(ws => RegexRtos.IsMatch(ws.Name));
                    break;
                case EnumBlock.Dvdc:
                    sheets = sheets.Where(ws => ws.Name.StartsWith(NeededSheets.PrefixWireless, StringComparison.OrdinalIgnoreCase) ||
                    ws.Name.StartsWith(NeededSheets.PrefixLcd, StringComparison.OrdinalIgnoreCase) ||
                    ws.Name.StartsWith(HardIpConstData.PrefixDcTest, StringComparison.OrdinalIgnoreCase));
                    break;
            }
            var hardIpReader = new TestPlanConverter(ScghStatic.ScghData);
            result.PlanDic = hardIpReader.ReadSheet(sheets.ToList(), null);

            return result;
        }
    }
}
