using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.Reader.ConfigFile.RtosTable;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;
namespace Automation.GenerateIgxl.SpiRom
{
    internal class GenSpiRom
    {
        public const string SpiromCodeFileSheetName = "SpiromCodeFile";
        private readonly ExcelWorkbook _excelWorkbook;
        private readonly SpiRomOutPut _outPutData = new SpiRomOutPut();

        public GenSpiRom(ExcelWorkbook excelWorkbook)
        {
            _excelWorkbook = excelWorkbook;
        }

        public bool Workflow()
        {

            ExcelWorkbook targetWb = _excelWorkbook;
            List<ExcelWorksheet> spiRomSheets =
                EpWorkbook.TestPlanWorkbook.Worksheets
                .Where(s => SearchInfo.IsSpiRomSheet(s.Name))
                .ToList();

            foreach (ExcelWorksheet sheet in spiRomSheets)
            {
                ExcelWorksheet existing = targetWb.Worksheets[sheet.Name];

                if (existing != null)
                {
                    targetWb.Worksheets.Delete(existing);
                }

                targetWb.Worksheets.Add(sheet.Name, sheet);
            }


            foreach (ExcelWorksheet worksheet in targetWb.Worksheets)
            {
                if (worksheet.Name == SpiromCodeFileSheetName)
                {
                    CreateSpiromCodeFileSheet(worksheet);
                }
            }
            var types = new List<EnumSheetType> { EnumSheetType.DTFlowtableSheet, EnumSheetType.DTBintablesSheet, EnumSheetType.DTTestInstancesSheet, EnumSheetType.DTTestInstancesSheet, EnumSheetType.DTLevelSheet, EnumSheetType.DTPatternSetSheet, EnumSheetType.DTTimesetBasicSheet };
            var igxlDataLoader = new IgxlLoader(targetWb, types);

            _outPutData.LevelList = igxlDataLoader.LevelSheets;
            _outPutData.PatSetlList = igxlDataLoader.PatSetSheets;
            _outPutData.TimeSetList = igxlDataLoader.TimeSetBasicSheets;

            string instSheetName = "TestInst_SPIROM_CS";
            List<string> mainFlowName = new List<string> { "Flow_Table_Write_SPIROM_main_CS" };


            foreach (SubFlowSheet flowSheet in igxlDataLoader.FlowSheets)
            {
                if (mainFlowName.Exists(x => x.Equals(flowSheet.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    CreateFlowSheet(flowSheet);
                }
            }

            foreach (InstanceSheet instanceSheet in igxlDataLoader.InstanceSheets)
            {
                if (instanceSheet.Name.Equals(instSheetName))
                {
                    CreateInstanceSheet(instanceSheet);
                }
            }

            CreateBinTableRows(_outPutData.FlowList);

            _outPutData.OutPutNonIgxlToTxt(FolderStructure.DirSpiRom, SpiromCodeFileSheetName);
            _outPutData.AddDataToLocalSpace();

            return true;
        }

        private void CreateBinTableRows(List<SubFlowSheet> flowList)
        {
            List<BinTableRow> binTableRows = new List<BinTableRow>();
            foreach (SubFlowSheet flow in flowList)
            {
                foreach (FlowRow flowRow in flow.Rows)
                {
                    if (string.Equals(flowRow.Opcode, "binTable", StringComparison.OrdinalIgnoreCase))
                    {
                        var binTableRow = new BinTableRow { Name = flowRow.Parameter, ItemList = flowRow.Parameter.Replace("Bin", "F") };
                        List<string> binNameCategory = flowRow.Parameter.Split('_').ToList();
                        string category1 = binNameCategory.Count >= 1 ? binNameCategory[1] : "";
                        string category2 = binNameCategory.Count >= 2 ? binNameCategory[2] : "";
                        binTableRow.Items.Add("T");
                        binTableRow.Op = "AND";
                        BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("SPIROM", category1, category2, binTableRow);
                        binTableRow.Sort = binNumInfo.SoftBin.ToString("G15");
                        binTableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                        binTableRow.Result = binNumInfo.BinNumInfo.Status;

                        binTableRows.Add(binTableRow);

                    }
                }
            }
            _outPutData.BinTableRows = binTableRows.GroupBy(x => x.Name).Select(x => x.First()).ToList();
        }

        private void CreateSpiromCodeFileSheet(ExcelWorksheet excelWorksheet)
        {
            string strBinFileName = "";
            ExcelWorksheet rtosTableSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(x => x.Name.Equals("Rtos_Table", StringComparison.OrdinalIgnoreCase));
            if (rtosTableSheet != null)
            {
                strBinFileName = RtosTableSheet.LoadConfig(rtosTableSheet).BinFileName;
            }

            StringBuilder lSpiromCodeFileStringBuilder = new StringBuilder();
            for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++)
                {
                    if (!string.IsNullOrEmpty(strBinFileName) && (i == 1 || i == 2) && j == 2)
                    {
                        lSpiromCodeFileStringBuilder.Append(@".\PATTERN\SPIROM\" + strBinFileName);
                    }
                    else
                    {
                        lSpiromCodeFileStringBuilder.Append(excelWorksheet.GetCellValue(i, j));
                    }

                    lSpiromCodeFileStringBuilder.Append(CommonConst.Tab);
                }
                lSpiromCodeFileStringBuilder.Append(CommonConst.Enter);
            }
            _outPutData.SpiromCodeFileString = lSpiromCodeFileStringBuilder.ToString();
        }

        private void CreateFlowSheet(SubFlowSheet subFlowSheet)
        {
            subFlowSheet.Name = subFlowSheet.Name.Replace("_CS", "");
            subFlowSheet.SourceInfo.Name = subFlowSheet.Name;
            _outPutData.FlowList.Add(subFlowSheet);
        }

        private void CreateInstanceSheet(InstanceSheet instanceSheet)
        {
            instanceSheet.Name = instanceSheet.Name.Replace("_CS", "");
            foreach (InstanceRow instance in instanceSheet.Rows)
            {
                if (!instance.VbtType.Equals(".NET", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                instance.VbtName = TestProgram.VbtFunctionLib.GetFunctionByName(instance.VbtName.Split('.').LastOrDefault(), "").FullFunctionName;
            }
            _outPutData.InstanceList.Add(instanceSheet);
        }
    }
}
