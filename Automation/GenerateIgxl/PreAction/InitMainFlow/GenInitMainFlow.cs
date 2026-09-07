using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.PreAction.InitMainFlow
{
    public class GenInitMainFlow
    {
        private const string ConFlowTableMainInitVar = "Flow_Table_Main_Init_Var";
        private readonly List<SubFlowSheet> _initSubFlowSheets = new List<SubFlowSheet>();

        public void WorkFlow(bool pFlag)
        {

            if (!pFlag)
            {
                return;
            }

            ExcelWorkbook workbook = SettingStatic.BasicConfigWorkbook;
            const string lStrPattern = @"^Flow_Table_Main_Init\w*";
            foreach (ExcelWorksheet workSheet in workbook.Worksheets)
            {
                if (!Regex.IsMatch(workSheet.Name, lStrPattern, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(workSheet.Name, ConFlowTableMainInitVar, RegexOptions.IgnoreCase))
                {
                    continue;
                }

                if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder)
                    && workSheet.Name.Equals("Flow_Table_Main_Init_Flows_Cs", StringComparison.OrdinalIgnoreCase))
                {
                    continue; //By pass Flow_Table_Main_Init_Flows_Cs when using VBT library.
                }

                if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder)
                    && workSheet.Name.Equals("Flow_Table_Main_Init_Flows", StringComparison.OrdinalIgnoreCase))
                {
                    continue; //By pass Flow_Table_Main_Init_Flows when using C# library.
                }

                if (workSheet.Name.Equals("Flow_Table_Main_Init_EnableWd", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                var flowReader = new ReadFlowSheet();
                SubFlowSheet flow = flowReader.ReadSheet(workSheet);
                if (flow.Name.Equals("Flow_Table_Main_Init_Flows_Cs", StringComparison.OrdinalIgnoreCase))
                {
                    flow.Name = "Flow_Table_Main_Init_Flows";
                }
                _initSubFlowSheets.Add(flow);
            }

            AddFlowToLocalSpecs();

            AddBinTable();
        }

        private void AddFlowToLocalSpecs()
        {
            List<string> flowNameList = TestProgram.IgxlWorkBk.GetFlowSheetNameList();
            foreach (SubFlowSheet sheet in _initSubFlowSheets)
            {
                sheet.SourceInfo.Name = "Initialize";
                if (flowNameList.Contains(sheet.Name))
                {
                    throw new Exception($"Flow {sheet} has been added!");
                }
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirMain, sheet);
            }
        }

        private void AddBinTable()
        {
            BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            var binTableRow = new BinTableRow { Name = "Bin_EFUSE_ecid_other", ItemList = "F_ecid_other", Op = "AND" };
            binTableRow.Items.Add("T");

            BinNumResult binInfo = GetBinInfo(binTableRow);

            binTableRow.Sort = binInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binInfo.BinNumInfo.Status;
            binTable.AddRow(binTableRow);

            binTableRow = new BinTableRow { Name = "Bin_initFlow_BinOut", ItemList = "F_initFlow_BinOut", Op = "AND" };
            binTableRow.Items.Add("T");
            binInfo = GetBinInfo(binTableRow);
            binTableRow.Sort = binInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binInfo.BinNumInfo.Status;
            binTable.AddRow(binTableRow);

            binTableRow = new BinTableRow { Name = "Bin_OCR_error", ItemList = "F_OCR_error", Op = "AND" };
            binTableRow.Items.Add("T");
            binInfo = GetBinInfo(binTableRow);
            binTableRow.Sort = binInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binInfo.BinNumInfo.Status;
            binTable.AddRow(binTableRow);

            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(new List<string> { "F_ecid_other" }, EFuseConst.Efuse, FolderStructure.DirMain);
        }

        private BinNumResult GetBinInfo(BinTableRow binTableRow)
        {
            string[] binNameSegments = binTableRow.Name.Split('_');
            string category1 = binNameSegments.Length > 1 ? binNameSegments[1] : "";
            string category2 = binNameSegments.Length > 2 ? binNameSegments[2] : "";
            return BinNumberSingleton.Instance.GetBinInfo("InitFlow", category1, category2, binTableRow);
        }
    }
}
