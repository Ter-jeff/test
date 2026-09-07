using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using CommonLib.Extension;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using RfLib.Dvdc.Reader.CapturePostProcess;
using RfLib.Dvdc.Reader.DsscSetup;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public class RegisterTableGenerator(ExcelPackage excelPackage)
    {
        public ExcelPackage XlPackage = excelPackage;
        public ExcelWorkbook XlWorkBook = excelPackage.Workbook;

        public void GenRegisterTable(Dictionary<EnumVbtFuncType, DsscSetupSheet> registerSheetList)
        {

            foreach (KeyValuePair<EnumVbtFuncType, DsscSetupSheet> registerSheet in registerSheetList)
            {
                var mergeRows = new List<DsscSetupSheetRow>();
                ExcelWorksheet excelWorksheet;

                if (registerSheet.Key == EnumVbtFuncType.LCD)
                {
                    excelWorksheet = XlWorkBook.AddSheet("DSSCSetupSheet_Meas");
                }
                else
                {
                    excelWorksheet = XlWorkBook.AddSheet("DSSCSetupSheet_" + registerSheet.Key);
                }

                mergeRows.AddRange(registerSheet.Value.InitRowList);
                mergeRows.AddRange(registerSheet.Value.RowList);

                MemberInfo[] skipBlock = [.. typeof(DsscSetupSheetRow).GetProperties().Where(p => p.Name != "Block" && p.Name != "Type").Select(p => (MemberInfo)p)];

                if (mergeRows.Count > 0)
                {
                    excelWorksheet.Cells.LoadFromCollection(mergeRows, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, skipBlock);
                    //ComEpplusExcel.Instance().MergeCell(excelWorksheet, 1);
                    excelWorksheet.Cells.TryAutoFitColumns();
                }

                //post process
                foreach (KeyValuePair<string, List<PostProcessSheetRow>> dic in registerSheet.Value.PostProcessRowList.GroupBy(x => x.BlockName).ToDictionary(x => x.Key, x => x.ToList()))
                {
                    var rows = new List<PostProcessSheetRow>();
                    rows.AddRange(dic.Value);
                    excelWorksheet = XlWorkBook.AddSheet("CPP_" + dic.Key);

                    MemberInfo[] headers = [.. typeof(PostProcessSheetRow).GetProperties().Where(p => p.Name != "BlockName").Select(p => (MemberInfo)p)];

                    int startCol = 1;
                    int startRow = 1;
                    int offset = 0;

                    startRow++;
                    foreach (KeyValuePair<string, List<PostProcessSheetRow>> setup in dic.Value.GroupBy(x => x.SetupName).ToDictionary(x => x.Key, x => x.ToList()))
                    {
                        foreach (MemberInfo row in headers)
                        {
                            excelWorksheet.Cells[1, startCol].Value = row.Name;
                            startCol++;
                        }
                        bool isFirst = true;
                        startRow = 2;
                        foreach (PostProcessSheetRow row in setup.Value)
                        {
                            if (isFirst)
                            {
                                excelWorksheet.Cells[startRow, PostProcessSheetRow.SetupNameIdx + offset].Value = row.SetupName;
                                excelWorksheet.Cells[startRow, PostProcessSheetRow.PatternNameIdx + offset].Value = row.PatternName;

                                /**
                                 * SIMULATIONFILE:
                                 * FORCEFLOWFLAG:True
                                 * CAPTUREDATAPRINT:False
                                 */
                                excelWorksheet.Cells[startRow + 1, PostProcessSheetRow.SetupNameIdx + offset].Value = "SIMULATIONFILE:";
                                excelWorksheet.Cells[startRow + 2, PostProcessSheetRow.SetupNameIdx + offset].Value = "FORCEFLOWFLAG:True";
                                excelWorksheet.Cells[startRow + 3, PostProcessSheetRow.SetupNameIdx + offset].Value = "CAPTUREDATAPRINT:False";
                                isFirst = false;
                            }
                            excelWorksheet.Cells[startRow, PostProcessSheetRow.TestNameIdx + offset].Value = row.TestName;
                            excelWorksheet.Cells[startRow, PostProcessSheetRow.BitWidthIdx + offset].Value = row.BitWidth;
                            excelWorksheet.Cells[startRow, PostProcessSheetRow.StoreNameIdx + offset].Value = row.StoreName;
                            startRow++;
                        }
                        offset += 10;
                    }
                    excelWorksheet.Cells.TryAutoFitColumns();
                }
            }
        }

    }
}
