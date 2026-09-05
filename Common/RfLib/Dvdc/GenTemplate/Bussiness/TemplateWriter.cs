using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;
using RfLib.Dvdc.Reader.DsscSetup;
using RfLib.InstrumentSetup;

using DataTable = System.Data.DataTable;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public class TemplateWriter
    {
        public string Filename;
        public Dictionary<EnumVbtFuncType, DsscSetupSheet> RegisterTable { get; set; } = [];
        public DataTable? InstrumentSetupTable { get; set; }
        public List<SetEfuseItem> FuseInfos { get; set; } = [];
        public Dictionary<string, List<string>> MissingReg { get; set; } = [];
        public Dictionary<string, List<string>> DupRegName { get; set; } = [];
        public HardIpInputData HardIpInputData { get; }

        public TemplateWriter(HardIpInputData hardIpInputData)
        {
            if (!Directory.Exists(LocalSpecs.TarFolder))
            {
                Directory.CreateDirectory(LocalSpecs.TarFolder);
            }

            Filename = string.Format("{0}\\{1}_TestPlanTemplate{2}.xlsx", LocalSpecs.TarFolder, LocalSpecs.CurrentProject, VersionControl.Timestamp);
            HardIpInputData = hardIpInputData;
            //if (!LocalSpecs.TestPlanFileName.Equals("N/A"))
            //{
            //    File.Copy(LocalSpecs.TestPlanFileName, _filename);
            //}
        }

        public void WriteTemplate(Dictionary<string, List<TemplateRow>> templates)
        {
            using var excel = new ExcelPackage(new FileInfo(Filename));
            try
            {
                foreach (string sheetName in templates.Keys)
                {
                    List<TemplateRow> templateRows = templates[sheetName];
                    //var existWireless = templates[sheetName].OfType<WirelessTemplateRow>().ToList().Count > 0
                    //    && templateRows.OfType<WirelessTemplateRow>().Any(p =>
                    //        !string.IsNullOrEmpty(p.InstrumentSetup) ||
                    //        !string.IsNullOrEmpty(p.TrimMeas) ||
                    //        p.TrimRegName.Count>0 ||
                    //        !string.IsNullOrEmpty(p.TrimTarget) ||
                    //        !string.IsNullOrEmpty(p.TrimType) ||
                    //        !string.IsNullOrEmpty(p.BestCode) ||
                    //        !string.IsNullOrEmpty(p.PostCalc));
                    //||!string.IsNullOrEmpty(p.Interpose));

                    string worksheetName = sheetName;
                    if (LocalSpecs.Options.Device == EnumDevice.LCD)
                    {
                        //continue;
                        worksheetName = worksheetName.Replace("HardIP", "LCD");
                    }

                    if (LocalSpecs.Options.Device == EnumDevice.RF)
                    {
                        worksheetName = worksheetName.Replace("HardIP", "Wireless");
                    }
                    else if (worksheetName.EqualsIgnoreCase("HardIP_IDS"))
                    {
                        worksheetName = "DCTEST_IDS";
                    }

                    ExcelWorksheet? wSheet = null;
                    wSheet = excel.Workbook.Worksheets[worksheetName] ??
                             excel.Workbook.Worksheets.Add(worksheetName);
                    Template template = new Template();
                    #region Write Headers
                    if (LocalSpecs.Options.Device == EnumDevice.RF ||
                        LocalSpecs.Options.Device == EnumDevice.LCD)
                    {
                        template = new WirelessTemplate();
                    }
                    else //Else use HardIP
                    {
                        template = new Template();
                    }
                    template.WriteTemplateHeader(wSheet);
                    #endregion

                    #region Write Content
                    int rowIndex = 2;

                    #endregion
                    Dictionary<int, List<TemplateRow>>.ValueCollection testItems = templateRows.GroupBy(p => p.TestItem)
                        .ToDictionary(p => p.Key, p => p.ToList()).Values;
                    foreach (List<TemplateRow> testitem in testItems)
                    {
                        try
                        {
                            foreach (TemplateRow row in testitem)
                            {
                                row.Rowindex = rowIndex;
                                template.WriteTemplateContent(wSheet, row, rowIndex);
                                rowIndex++;
                            }
                            wSheet.Column(template.MiscInfo.Index).TryAutoFit();
                        }
                        catch (Exception ex)
                        {
                            ErrorMessageBox.Show(string.Format(ex.ToString()));
                        }
                        if (LocalSpecs.Options.Device == EnumDevice.RF)
                        {
                            var temp = (WirelessTemplate)template;
                            MergeRfInstSetupCell(wSheet, testitem, temp.RfInstrumentSetup.Index);
                            MergeRfCalcCell(wSheet, testitem, temp.RfTestType.Index);
                        }
                        MergeForceCell(wSheet, testitem, template.ForceCondition.Index);
                        MergeTestItemCell(wSheet, testitem, template.TestItem.Index);

                    }

                    //CellMerge(wSheet, 2);

                    for (int i = 1; i <= wSheet.Dimension.Columns; i++)
                    {
                        wSheet.Column(i).Width = 15;
                    }

                    wSheet.Row(1).Height = 40;
                    if (LocalSpecs.Options.Device == EnumDevice.LCD ||
                        LocalSpecs.Options.Device == EnumDevice.RF)
                    {
                        foreach (int index in new WirelessTemplate().GetNonUseIndex())
                        {
                            wSheet.Column(index).Hidden = true;
                            wSheet.Column(index).Width = 0;

                        }

                        //wSheet.Cells.AutoFitColumns();
                    }
                }

                var registerTableGenerator = new RegisterTableGenerator(excel);
                registerTableGenerator.GenRegisterTable(RegisterTable);
                if (InstrumentSetupTable != null)
                {
                    var instrumentSetupGenerator = new InstrumentSetupGenerator(excel);
                    instrumentSetupGenerator.GenInstrumentSetupTable(InstrumentSetupTable);
                }

                if (FuseInfos.Count > 0)
                {
                    ExcelWorksheet ws = excel.Workbook.Worksheets.Add("CaptoEfuseSetup");
                    WriteFuseInfo(ws, FuseInfos);
                }

                if (MissingReg != null)
                {
                    var errorcheckGenerator = new ErrorCheckerGenerator(excel);
                    errorcheckGenerator.GenMissingReg(MissingReg);

                }
                if (DupRegName != null)
                {
                    var errorcheckGenerator = new ErrorCheckerGenerator(excel);
                    errorcheckGenerator.GendupRegName(DupRegName);
                }

                if (HardIpInputData.HardIpRegAssigns.Count > 0)
                {
                    new RegAssignGenerator(HardIpInputData).WorkFlow(excel.Workbook, HardIpInputData.HardIpRegAssigns);
                }

                excel.Save();

                #region todo
                if (LocalSpecs.TestPlanFileName != "N/A")
                {
                    //using (var excel_new = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName)))
                    //{

                    //    foreach (var sh in excel.Workbook.Worksheets)
                    //    {
                    //        try
                    //        {
                    //            var newws = excel_new.Workbook.Worksheets.Add(sh.Name, sh);

                    //        }
                    //        catch (Exception e)
                    //        {
                    //            ;
                    //        }
                    //    }
                    //    excel_new.File = new FileInfo(@"D:\Raze\test1.xlsx");
                    //    excel_new.Save();
                    //}

                }
                #endregion
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                throw;
            }

        }

        private static void WriteFuseInfo(ExcelWorksheet excelWorksheet, List<SetEfuseItem> setEfuseItems)
        {
            int rowindex = 1;
            foreach (SetEfuseItem setEfuseItem in setEfuseItems)
            {
                excelWorksheet.Cells[rowindex++, 1].Value = setEfuseItem.ItemName;
                for (int i = 1; i <= setEfuseItem.Headers.Count; i++)
                {
                    excelWorksheet.Cells[rowindex, i].Value = setEfuseItem.Headers[i - 1];
                }
                rowindex++;
                foreach (EFuseStoreRow fuseInfo in setEfuseItem.Contents)
                {
                    excelWorksheet.Cells[rowindex, 1].Value = fuseInfo.CaptureStoreName;
                    excelWorksheet.Cells[rowindex, 2].Value = fuseInfo.FieldName;
                    excelWorksheet.Cells[rowindex, 3].Value = fuseInfo.FuseEnable;
                    rowindex++;
                }
            }
        }

        private static void MergeRfInstSetupCell(ExcelWorksheet excelWorksheet, IEnumerable<TemplateRow> templateRows, int column)
        {
            int index_min = -1;
            int index_max = -1;

            var testRows = templateRows.OfType<WirelessTemplateRow>().ToList();
            var forceMergeItems = testRows.Where(p => !string.IsNullOrEmpty(p.InstrumentSetup)).GroupBy(p => p.InstrumentSetup).ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<WirelessTemplateRow>> forceMergeItem in forceMergeItems)
            {
                index_min = forceMergeItem.Value.Min(p => p.Rowindex);
                index_max = forceMergeItem.Value.Max(p => p.Rowindex);
                try
                {
                    if (excelWorksheet.Cells[index_min, column, index_max, column].Merge)
                    {
                        continue;
                    }

                    excelWorksheet.Cells[index_min, column, index_max, column].Merge = true;
                    excelWorksheet.Cells[index_min, column, index_max, column].Style.VerticalAlignment =
                        ExcelVerticalAlignment.Center;
                    excelWorksheet.Cells[index_min, column, index_max, column].Style.HorizontalAlignment =
                        ExcelHorizontalAlignment.Center;
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
        }

        private static void MergeRfCalcCell(ExcelWorksheet excelWorksheet, IEnumerable<TemplateRow> templateRows, int column)
        {
            int index_min = -1;
            int index_max = -1;

            var testRows = templateRows.OfType<WirelessTemplateRow>().ToList();
            var forceMergeItems = testRows.Where(p => !string.IsNullOrEmpty(p.InstrumentSetup)).GroupBy(p => p.InstrumentSetup).ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<WirelessTemplateRow>> forceMergeItem in forceMergeItems)
            {
                index_min = forceMergeItem.Value.Min(p => p.Rowindex);
                index_max = forceMergeItem.Value.Max(p => p.Rowindex);
                try
                {
                    if (excelWorksheet.Cells[index_min, column, index_max, column].Merge)
                    {
                        continue;
                    }

                    excelWorksheet.Cells[index_min, column, index_max, column].Merge = true;
                    excelWorksheet.Cells[index_min, column, index_max, column].Style.VerticalAlignment =
                        ExcelVerticalAlignment.Center;
                    excelWorksheet.Cells[index_min, column, index_max, column].Style.HorizontalAlignment =
                        ExcelHorizontalAlignment.Center;
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
        }

        private static void MergeForceCell(ExcelWorksheet excelWorksheet, IEnumerable<TemplateRow> templateRows, int column)
        {
            int index_min = -1;
            int index_max = -1;
            var forceMergeItems =
                templateRows.Where(p => !string.IsNullOrEmpty(p.Seqindex))
                    .GroupBy(p => p.Seqindex)
                    .ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<TemplateRow>> forceMergeItem in forceMergeItems)
            {
                index_min = forceMergeItem.Value.Min(p => p.Rowindex);
                index_max = forceMergeItem.Value.Max(p => p.Rowindex);
                excelWorksheet.Cells[index_min, column, index_max, column].Merge = true;
                excelWorksheet.Cells[index_min, column, index_max, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelWorksheet.Cells[index_min, column, index_max, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }

        private static void MergeTestItemCell(ExcelWorksheet excelWorksheet, List<TemplateRow> templateRows, int column)
        {
            int index_min = -1;
            int index_max = -1;
            index_min = templateRows.Min(p => p.Rowindex);
            index_max = templateRows.Max(p => p.Rowindex);
            excelWorksheet.Cells[index_min, column, index_max, column].Merge = true;
            excelWorksheet.Cells[index_min, column, index_max, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            excelWorksheet.Cells[index_min, column, index_max, column].Style.HorizontalAlignment =
                ExcelHorizontalAlignment.Center;
        }

    }
}
