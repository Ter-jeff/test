using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using EfuseCheckCmdLib.BDF;
using EfuseCheckCmdLib.CFGTable;
using EfuseCheckCmdLib.Datalog;
using EfuseCheckCmdLib.DataStructure;
using EfuseCheckCmdLib.EFuse.EFuseApp;

using OfficeOpenXml;

using TestPlanLib.Efuse;

using ValueType = EfuseCheckCmdLib.Base.ValueType;

namespace EfuseCheckCmdLib.Output
{
    public static partial class WriteExcel
    {
        private const string SetWriteVariableSheetName = "SetWriteVariable";
        public static readonly Regex RegBin = BinRegex();
        public static readonly Regex RegHex = HexRegex();

        [GeneratedRegex("^0*[bB](?<Number>[01]+)$")]
        private static partial Regex BinRegex();

        [GeneratedRegex("^0*[xX](?<Number>[0-9A-Fa-f]+)$")]
        private static partial Regex HexRegex();

        [GeneratedRegex("bank_", RegexOptions.IgnoreCase)]
        private static partial Regex BankPrefixRegex();

        [GeneratedRegex(@"(?<bank>\w*)\[")]
        private static partial Regex BankBracketRegex();

        [GeneratedRegex(@"\[(?<fieldName>.*)\]", RegexOptions.IgnoreCase)]
        private static partial Regex FieldNameBracketRegex();

        [GeneratedRegex("^=", RegexOptions.IgnoreCase)]
        private static partial Regex FormulaPrefixRegex();

        public static void WriteReport(List<DiceInfo> diceInfos, LoaderEfuseBitDef loaderEfuseBitDef, EfuseDramTable efuseDramTable, FuseCheckTable fuseCheckTable, string cfgPath, string outPutFile)
        {
            string filePath = outPutFile;
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            bool isContainsSetWrite = diceInfos.Any(x => x.SetWriteVariableLines != null && x.SetWriteVariableLines.Count != 0);

            using var package = new ExcelPackage(new FileInfo(filePath));
            int colIdxMain = PrintEfuseBitDefTable(diceInfos, loaderEfuseBitDef, package);

            colIdxMain = PrintPrr(diceInfos, package, colIdxMain);

            PrintConfigTable(diceInfos, loaderEfuseBitDef, cfgPath, package);

            PrintDram(diceInfos, efuseDramTable, package);

            WhiteFuseCheckTable(diceInfos, fuseCheckTable, package);

            SetWriteVariable(diceInfos, isContainsSetWrite, package);

            package.Save();
        }

        private static void SetWriteVariable(List<DiceInfo> diceInfos, bool isContainsSetWrite, ExcelPackage excelPackage)
        {
            #region SetWriteVariable
            if (isContainsSetWrite)
            {
                int colIndex = 3;
                ExcelWorksheet setWriteVariableSheet = excelPackage.Workbook.Worksheets.Add(SetWriteVariableSheetName);
                setWriteVariableSheet.Cells[2, 1].Value = "Dictionary Name";
                setWriteVariableSheet.Cells[2, 2].Value = "Block";
                foreach (DiceInfo diceInfo in diceInfos)
                {
                    int rowIndex = 1;
                    setWriteVariableSheet.Cells[1, colIndex].Value = $"X{diceInfo.XCoor}_Y{diceInfo.YCoor}";
                    foreach (SetWriteVariableLine setWriteVariable in diceInfo.SetWriteVariableLines)
                    {
                        setWriteVariableSheet.Cells[2, colIndex].Value = "Fuse Real Value";
                        setWriteVariableSheet.Cells[2, colIndex + 1].Value = "Meas Value";
                        EfuseRow row = diceInfo.EfuseRows.Find(x => x.SubConfig.EqualsIgnoreCase(setWriteVariable.Key) && x.Site == diceInfo.Site) ?? new EfuseRow
                        {
                            SubConfig =
                            $"NotExistsinEFuseReadWrite_{setWriteVariable.Key}",
                            BankConfig = "NotExistsinEFuseReadWrite"
                        };

                        string[]? rowLineArray = row.Line?.Line?.Split([' ', '=', '[', ']', '\t'], StringSplitOptions.RemoveEmptyEntries);
                        string extractedRealValue = rowLineArray?.Length >= 2 ? rowLineArray[^2] : "N/A";

                        setWriteVariableSheet.Cells[2 + rowIndex, 1].Value = row.SubConfig;
                        setWriteVariableSheet.Cells[2 + rowIndex, 2].Value = row.BankConfig;
                        setWriteVariableSheet.Cells[2 + rowIndex, colIndex].Value = extractedRealValue;
                        setWriteVariableSheet.Cells[2 + rowIndex, colIndex + 1].Value = setWriteVariable.Value;
                        if (int.TryParse(extractedRealValue, out int extractedInt) && extractedRealValue != setWriteVariable.Value)
                        {
                            ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_01, setWriteVariableSheet.Name, rowIndex + 2, colIndex + 1, [extractedRealValue, setWriteVariable.Value, extractedRealValue]);
                            HighLightCell(setWriteVariableSheet, rowIndex + 2, colIndex + 1, rowIndex + 2, colIndex + 1, Color.Red);
                        }
                        rowIndex++;
                    }
                    colIndex += 2;
                }
                setWriteVariableSheet.Cells.TryAutoFitColumns();
            }

            #endregion
        }

        private static void WhiteFuseCheckTable(List<DiceInfo> diceInfos, FuseCheckTable fuseCheckTable, ExcelPackage excelPackage)
        {
            #region FuseCheck Table
            if (fuseCheckTable.InPath != null)
            {
                ExcelWorksheet fuseworksheet = excelPackage.Workbook.Worksheets.Add("FuseCheckTable");
                int shiftcolIdx = 1;
                int shiftrowIdx = 0;

                WriteData(diceInfos, fuseCheckTable, fuseworksheet, ref shiftcolIdx, ref shiftrowIdx);

                if (fuseworksheet.Cells.Value != null)
                {
                    for (int col = 1; col <= fuseworksheet.Dimension.End.Column; col++)
                    {
                        ExcelRange cell = fuseworksheet.Cells[1, col];
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                }
                else
                {
                    fuseworksheet.Cells[1, 1].Value = "No matching bin results with FuseCheck table";
                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_02, fuseworksheet.Name, 1, 1, []);
                    HighLightCell(fuseworksheet, 1, 1, 1, 1, Color.Red);
                }
                fuseworksheet.Cells.TryAutoFitColumns();
            }
            #endregion
        }

        private static void WriteData(List<DiceInfo> diceInfos, FuseCheckTable fuseCheckTable, ExcelWorksheet excelWorksheet, ref int shiftcolIdx, ref int shiftrowIdx)
        {
            foreach (DiceInfo diceInfo in diceInfos)
            {
                List<EfuseRow> efuseRows = diceInfo.EfuseRows;
                int efusebin = diceInfo.SortBin;
                List<int> indices = [];
                for (int i = 2; i < fuseCheckTable.Titles.Count; i++)
                {
                    string titleVal = "";
                    if (fuseCheckTable.Titles[i].Split('_').Length > 1)
                    {
                        titleVal = fuseCheckTable.Titles[i].Split('_')[0];
                    }
                    else
                    {
                        titleVal = fuseCheckTable.Titles[i];
                    }
                    if (int.TryParse(titleVal, out int number))
                    {
                        if (efusebin == number)
                        {
                            indices.Add(i);
                        }
                    }
                }
                if (indices.Count > 0)
                {
                    try
                    {
                        excelWorksheet.Cells[1, 1].Value = "Bank";
                        excelWorksheet.Cells[1, 2].Value = "Field";

                        foreach (int i in indices) // index of EfuseCheck table
                        {
                            excelWorksheet.Cells[1, 2 + shiftcolIdx].Value = fuseCheckTable.Titles[i];
                            excelWorksheet.Cells[1, 2 + shiftcolIdx + 1].Value =
                                $"X{diceInfo.XCoor}_Y{diceInfo.YCoor} ({diceInfo.SortBin})";
                            shiftrowIdx = 1;

                            foreach (List<string> fCheck in fuseCheckTable.Rows)
                            {
                                fCheck[1] = fCheck[1].Split('[')[0];
                                string binValue = efuseRows.Find(x => x.BankConfig.EqualsIgnoreCase(fCheck[0]) && x.SubConfig.EqualsIgnoreCase(fCheck[1]))!.BinValue;
                                string hexBinValue = efuseRows.Find(x => x.BankConfig.EqualsIgnoreCase(fCheck[0]) && x.SubConfig.EqualsIgnoreCase(fCheck[1]))!.HexValue;
                                string tableValue = fCheck[i];
                                string realValue = "";
                                bool resultWalkingOne = true;
                                excelWorksheet.Cells[1 + shiftrowIdx, 1].Value = fCheck[0];
                                excelWorksheet.Cells[1 + shiftrowIdx, 2].Value = fCheck[1];
                                excelWorksheet.Cells[1 + shiftrowIdx, 2 + shiftcolIdx].Value = tableValue;

                                if (tableValue.Split('|').Length > 1)
                                {
                                    foreach (string s in tableValue.Split('|'))
                                    {
                                        string tVal = s.Trim();
                                        if (tVal.StartsWithIgnoreCase("b"))
                                        {
                                            tableValue = tVal;
                                            excelWorksheet.Cells[1 + shiftrowIdx, 2 + shiftcolIdx + 1].Value = binValue;
                                            realValue = binValue;
                                        }
                                        else if (tVal.StartsWithIgnoreCase("w"))
                                        {
                                            int tableWalkingOne = int.Parse(tVal.Split('-')[1]);
                                            int countWalkingOne = hexBinValue.Count(c => c == '1');
                                            if (tableWalkingOne != countWalkingOne)
                                            {
                                                resultWalkingOne = false;
                                            }

                                        }
                                        else
                                        {
                                            tableValue = tVal;
                                            excelWorksheet.Cells[1 + shiftrowIdx, 2 + shiftcolIdx + 1].Value = hexBinValue;
                                            realValue = hexBinValue;
                                        }
                                    }
                                    if (!EfuseCmdUtility.CheckEfuseDefaultValue(realValue, tableValue) && !resultWalkingOne)
                                    {
                                        ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_03, excelWorksheet.Name, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, [tableValue, realValue, tableValue]);
                                        HighLightCell(excelWorksheet, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, Color.Red);
                                    }
                                }
                                else
                                {
                                    if (tableValue.StartsWithIgnoreCase("b"))
                                    {
                                        excelWorksheet.Cells[1 + shiftrowIdx, 2 + shiftcolIdx + 1].Value = binValue;
                                        realValue = binValue;
                                    }
                                    else
                                    {
                                        if (tableValue.StartsWithIgnoreCase("w"))
                                        {
                                            int tableWalkingOne = int.Parse(tableValue.Split('-')[1]);
                                            int countWalkingOne = hexBinValue.Count(c => c == '1');
                                            if (tableWalkingOne != countWalkingOne)
                                            {
                                                resultWalkingOne = false;
                                            }
                                        }
                                        excelWorksheet.Cells[1 + shiftrowIdx, 2 + shiftcolIdx + 1].Value = hexBinValue;
                                        realValue = hexBinValue;
                                    }
                                    if (tableValue.StartsWithIgnoreCase("w") && !resultWalkingOne)
                                    {
                                        ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_04, excelWorksheet.Name, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, [tableValue, realValue, tableValue]);
                                        HighLightCell(excelWorksheet, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, Color.Red);
                                    }
                                    else if (!tableValue.StartsWithIgnoreCase("w"))
                                    {
                                        if (!EfuseCmdUtility.CheckEfuseDefaultValue(realValue, tableValue))
                                        {
                                            ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_05, excelWorksheet.Name, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, [tableValue, realValue, tableValue]);
                                            HighLightCell(excelWorksheet, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, Color.Red);
                                        }
                                    }
                                }
                                shiftrowIdx++;
                            }
                            shiftcolIdx += 2;
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_06, excelWorksheet.Name, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, []);
                        HighLightCell(excelWorksheet, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, 1 + shiftrowIdx, 2 + shiftcolIdx + 1, Color.Red);
                        Console.WriteLine(string.Format(ex.ToString()));
                    }
                }
            }
        }

        private static void PrintDram(List<DiceInfo> diceInfos, EfuseDramTable efuseDramTable, ExcelPackage excelPackage)
        {
            #region Print DRAM
            if (efuseDramTable.InPath != null)
            {
                ExcelWorksheet dRaMworksheet = excelPackage.Workbook.Worksheets.Add("DRAM_CONFIG");
                int colIdxDram = 0;
                int shiftcolIdx = 1;
                Type typenum = typeof(DramReport1);
                PropertyInfo[] fieldsNum = typenum.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                string dRamType = diceInfos.First().DramType;
                if (string.IsNullOrEmpty(dRamType))
                {
                    dRaMworksheet.Cells[1, 1].Value = "No such DRAM files in EnableWords";
                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_07, dRaMworksheet.Name, 1, 1, []);
                    HighLightCell(dRaMworksheet, 1, 1, 1, 1, Color.Red);
                }
                else
                {
                    bool isDramMatch = false;
                    foreach (string eDram in efuseDramTable.Titles)
                    {
                        if (eDram == dRamType.Trim())
                        {
                            isDramMatch = true;
                            break;
                        }
                        colIdxDram++;
                    }
                    if (isDramMatch)
                    {
                        var dRamReports = new List<DramReport1>();
                        foreach (List<string> eDram in efuseDramTable.Rows)
                        {
                            var dRamReport = new DramReport1
                            {
                                Category = eDram[0],
                                Bank = eDram[^1],
                                Field = eDram[1],
                                DramType = eDram[colIdxDram],
                            };
                            dRamReports.Add(dRamReport);
                        }

                        foreach (DiceInfo diceInfo in diceInfos)
                        {
                            List<EfuseRow> efuseRows = diceInfo.EfuseRows;
                            dRaMworksheet.Cells[1, fieldsNum.Length + shiftcolIdx].Value =
                                $"X{diceInfo.XCoor}_Y{diceInfo.YCoor}";
                            int shiftrowIdx = 1;
                            foreach (List<string> eDram in efuseDramTable.Rows)
                            {
                                string[] bankConfig = eDram[^1].Split('[');
                                string binValue = efuseRows.Find(x => x.BankConfig.EqualsIgnoreCase(bankConfig[0]) && x.SubConfig.EqualsIgnoreCase(eDram[1]))!.BinValue;
                                string hexBinValue = efuseRows.Find(x => x.BankConfig.EqualsIgnoreCase(bankConfig[0]) && x.SubConfig.EqualsIgnoreCase(eDram[1]))!.HexValue;
                                string tableValue = eDram[colIdxDram];
                                string realValue = "";
                                if (tableValue.StartsWithIgnoreCase("b"))
                                {
                                    dRaMworksheet.Cells[1 + shiftrowIdx, fieldsNum.Length + shiftcolIdx].Value = binValue;
                                    realValue = binValue;
                                }
                                else
                                {
                                    dRaMworksheet.Cells[1 + shiftrowIdx, fieldsNum.Length + shiftcolIdx].Value = hexBinValue;
                                    realValue = hexBinValue;
                                }

                                if (!EfuseCmdUtility.CheckEfuseDefaultValue(realValue, tableValue))
                                {
                                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_08, dRaMworksheet.Name, 1 + shiftrowIdx, fieldsNum.Length + shiftcolIdx, [eDram[colIdxDram], realValue, tableValue]);
                                    HighLightCell(dRaMworksheet, 1 + shiftrowIdx, fieldsNum.Length + shiftcolIdx, 1 + shiftrowIdx, fieldsNum.Length + shiftcolIdx, Color.Red);
                                }
                                shiftrowIdx++;
                            }
                            shiftcolIdx++;
                        }

                        dRaMworksheet.Cells[1, 1].LoadFromCollection(dRamReports, true);
                        dRaMworksheet.Cells[1, 4].Value = efuseDramTable.Titles[colIdxDram];
                    }
                    else
                    {
                        dRaMworksheet.Cells[1, 1].Value = "No matching DRAM type in EnableWords";
                        ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_09, dRaMworksheet.Name, 1, 1, []);
                        HighLightCell(dRaMworksheet, 1, 1, 1, 1, Color.Red);
                    }
                }
                if (dRaMworksheet.Cells.Value != null)
                {
                    for (int col = 1; col <= dRaMworksheet.Dimension.End.Column; col++)
                    {
                        ExcelRange cell = dRaMworksheet.Cells[1, col];
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                }
                dRaMworksheet.Cells.TryAutoFitColumns();
            }
            #endregion
        }

        private static void PrintConfigTable(List<DiceInfo> diceInfos, LoaderEfuseBitDef loaderEfuseBitDef, string cfgPath, ExcelPackage excelPackage)
        {
            #region Print Config_Table
            if (!string.IsNullOrEmpty(cfgPath))
            {
                Dictionary<string, EfuseCfgTable> tables = CfgTableReader.CfgTable;
                foreach (EfuseCfgTable table in tables.Values)
                {
                    int colIdxConfig = 1;
                    if (excelPackage.Workbook.Worksheets[table.TableName] == null)
                    {
                        ExcelWorksheet cfgSheet = excelPackage.Workbook.Worksheets.Add(table.TableName);
                        string scenario = diceInfos.First().Scenario;
                        PrintCfgTableHeader(cfgSheet, table, ref colIdxConfig, scenario);
                        colIdxConfig++;

                        foreach (DiceInfo dice in diceInfos)
                        {
                            EfuseCmdUtility.CheckCfgCondition(cfgSheet, dice, loaderEfuseBitDef.BitDefTable, table, colIdxConfig);
                            colIdxConfig++;
                        }
                    }
                }
            }
            #endregion
        }

        private static int PrintPrr(List<DiceInfo> diceInfos, ExcelPackage excelPackage, int colIdxMain)
        {
            #region Print PRR
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("PRR_Check");
            var prrReports = new List<PrrReport>();
            foreach (DiceInfo diceInfo in diceInfos)
            {
                List<PrrRow> prrRows = diceInfo.PrrRows;
                List<EfuseRow> efuseRows = diceInfo.EfuseRows;
                foreach (PrrRow prrRow in prrRows)
                {
                    if (prrRow.Type.Contains(','))
                    {
                        string[] types = [.. prrRow.Type
                            .Split(',', StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())];

                        string[] efusePrrs = [.. types.Select(type => EfuseCmdUtility.BinaryStringToHexString(GetPrr(efuseRows, type)))];

                        string[] logPrrs = prrRow.Prr.Split('_', StringSplitOptions.RemoveEmptyEntries);
                        int count = Math.Min(types.Length, Math.Min(efusePrrs.Length, logPrrs.Length));

                        for (int i = 0; i < count; i++)
                        {
                            bool itemPass = EfuseCmdUtility.AreHexValuesEqual(efusePrrs[i], logPrrs[i]);
                            PrrReport? report = prrReports.FirstOrDefault(x => x.Site == prrRow.Site && x.Type == types[i]);
                            if (report == null)
                            {
                                report = new PrrReport
                                {
                                    Site = prrRow.Site,
                                    Type = types[i]
                                };
                                prrReports.Add(report);
                            }

                            report.Line = prrRow.Line.LineNo + 1;
                            report.PrrInLog = logPrrs[i];
                            report.PrrByEfuse = efusePrrs[i];
                            report.Result = itemPass ? "Pass" : "Fail";
                        }
                        continue;
                    }

                    string prr = EfuseCmdUtility.BinaryStringToHexString(GetPrr(efuseRows, prrRow.Type));
                    bool isPass = EfuseCmdUtility.AreHexValuesEqual(prr, prrRow.Prr);
                    PrrReport? existReport = prrReports.FirstOrDefault(x => x.Site == prrRow.Site && x.Type == prrRow.Type);

                    if (existReport == null)
                    {
                        prrReports.Add(new PrrReport
                        {
                            Site = prrRow.Site,
                            Line = prrRow.Line.LineNo + 1,
                            Type = prrRow.Type,
                            PrrInLog = prrRow.Prr,
                            PrrByEfuse = prr,
                            Result = isPass ? "Pass" : "Fail"
                        });
                    }
                    else
                    {
                        existReport.Line = prrRow.Line.LineNo + 1;
                        existReport.PrrInLog = prrRow.Prr;
                        existReport.PrrByEfuse = prr;
                        existReport.Result = isPass ? "Pass" : "Fail";
                    }
                }
                colIdxMain++;
            }
            worksheet.Cells[1, 1].LoadFromCollection(prrReports, true);
            worksheet.Cells.TryAutoFitColumns();
            #endregion
            return colIdxMain;
        }

        private static int PrintEfuseBitDefTable(List<DiceInfo> diceInfos, LoaderEfuseBitDef loaderEfuseBitDef, ExcelPackage excelPackage)
        {
            #region Print EFUSE_BitDef_Table
            ExcelWorksheet bdfWorksheet = excelPackage.Workbook.Worksheets.Add("EFUSE_BitDef_Table");
            PrintHeader(loaderEfuseBitDef, bdfWorksheet);

            List<BitDefRow> bitDefRows = PrintItem(loaderEfuseBitDef, bdfWorksheet);

            int colIdxMain = PrintValue(diceInfos, loaderEfuseBitDef, bdfWorksheet, bitDefRows);

            bdfWorksheet.Column(1).Width = 30;
            bdfWorksheet.Column(loaderEfuseBitDef.BitDefTable.Titles.Count).Width = 43;
            #endregion
            return colIdxMain;
        }

        private static int PrintValue(List<DiceInfo> diceInfos, LoaderEfuseBitDef loaderEfuseBitDef, ExcelWorksheet excelWorksheet, List<BitDefRow> bitDefRows)
        {
            #region Print value from datalog
            List<BitDefRow> checkRows = [];
            int colIdxMain = loaderEfuseBitDef.BitDefTable.Titles.Count + 1;
            foreach (DiceInfo diceInfo in diceInfos)
            {
                excelWorksheet.Cells[1, colIdxMain].Value = diceInfo.Sort;
                excelWorksheet.Cells[2, colIdxMain].Value = diceInfo.SortBin;
                excelWorksheet.Cells[3, colIdxMain].Value = diceInfo.CurrentJobStage;
                excelWorksheet.Cells[4, colIdxMain].Value = $"X{diceInfo.XCoor}_Y{diceInfo.YCoor}";
                excelWorksheet.Cells[5, colIdxMain].Value = diceInfo.PrrCode;
                excelWorksheet.Cells[6, colIdxMain].Value = diceInfo.EFuseLotNumber;
                excelWorksheet.Cells[7, colIdxMain].Value = diceInfo.EFuseWaferId;
                if (!string.IsNullOrEmpty(diceInfo.EFuseDieX) && !string.IsNullOrEmpty(diceInfo.EFuseDieY))
                {
                    excelWorksheet.Cells[8, colIdxMain].Value = diceInfo.EFuseDieX;
                    excelWorksheet.Cells[9, colIdxMain].Value = diceInfo.EFuseDieY;
                }
                excelWorksheet.Cells[1, colIdxMain, 9, colIdxMain].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                foreach (EfuseRow efuseRow in diceInfo.EfuseRows)
                {
                    string efBank = BankPrefixRegex().Replace(efuseRow.BankConfig, "");
                    BitDefRow? bitDefRow = bitDefRows.FirstOrDefault(x => x.Name.EqualsIgnoreCase(efuseRow.SubConfig) && x.Block.Contains(efBank, StringComparison.OrdinalIgnoreCase));
                    if (bitDefRow != null)
                    {
                        checkRows.Add(bitDefRow);
                        string value = "";
                        ValueType type = bitDefRow.Type;
                        if (bitDefRow.IsCrc)
                        {
                            value = GetCrc(bitDefRow.Crc, loaderEfuseBitDef.BitDefTable, diceInfo);
                        }
                        else if (bitDefRow.Name.ContainsIgnoreCase("lot_id"))
                        {
                            value = efuseRow.Data;
                        }
                        else if (type == ValueType.Bin)
                        {
                            value = efuseRow.BinValue;
                        }
                        else if (type == ValueType.Dec)
                        {
                            value = efuseRow.DecValue;
                        }
                        else if (type == ValueType.Hex)
                        {
                            value = efuseRow.HexValue;
                        }
                        else if (type == ValueType.Real)
                        {
                            value = efuseRow.DecValue;
                        }
                        else
                        {
                            value = "";
                        }

                        if (bitDefRow.JobStage <= diceInfo.JobNum)
                        {
                            excelWorksheet.Cells[bitDefRow.RowNum, colIdxMain].Value = value;
                            if (bitDefRow.IsCrc)
                            {
                                string expected = efuseRow.HexValue;
                                if (!EfuseCmdUtility.AreHexValuesEqual(expected, value))
                                {
                                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_10, bitDefRow.SheetName, bitDefRow.RowNum, colIdxMain, [efuseRow.SubConfig, value, expected]);
                                    HighLightCell(excelWorksheet, bitDefRow.RowNum, colIdxMain, bitDefRow.RowNum, colIdxMain, Color.Red);
                                }
                            }
                            else if (bitDefRow.IsDefault)
                            {
                                if (!EfuseCmdUtility.CheckEfuseDefaultValue(value, bitDefRow.DefaultValue))
                                {
                                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_11, bitDefRow.SheetName, bitDefRow.RowNum, colIdxMain, [efuseRow.SubConfig, value, bitDefRow.DefaultValue]);
                                    HighLightCell(excelWorksheet, bitDefRow.RowNum, colIdxMain, bitDefRow.RowNum, colIdxMain, Color.Red);
                                }
                            }
                        }
                        else
                        {
                            #region Check item after programming stage
                            string logValue = EfuseCmdUtility.BinaryStringToDecimal(efuseRow.Bits);
                            excelWorksheet.Cells[bitDefRow.RowNum, colIdxMain].Value = logValue;
                            if (!double.TryParse(logValue, out double data) || data != 0.0)
                            {
                                ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_12, bitDefRow.SheetName, bitDefRow.RowNum, loaderEfuseBitDef.BitDefTable.PrgStageIdx + 1, [efuseRow.SubConfig, value, bitDefRow.DefaultValue]);
                                HighLightCell(excelWorksheet, bitDefRow.RowNum, loaderEfuseBitDef.BitDefTable.PrgStageIdx + 1, bitDefRow.RowNum, loaderEfuseBitDef.BitDefTable.PrgStageIdx + 1, Color.Red);
                            }
                            #endregion
                        }
                    }
                }

                //#region H/L limits check
                //if (diceInfo.LimitRows != null)
                //{
                //    foreach (LimitRow limitRow in diceInfo.LimitRows)
                //    {
                //        BitDefRow bitDefRow = bitDefRows.FirstOrDefault(x => x.Name.Equals(limitRow.TestName, StringComparison.CurrentCultureIgnoreCase));
                //        if (bitDefRow != null)
                //        {
                //            if (bitDefRow.LowLimit != "N/A")
                //            {
                //                double lo = EfuseCmdUtility.ToDec(bitDefRow.LowLimit);
                //                if (lo != limitRow.LowLimit)
                //                {
                //                    var error = new Error()
                //                    {
                //                        ErrorType = BinCutErrorType.E_Formula,
                //                        SheetName = bitDefRow.SheetName,
                //                        ErrorLevel = ErrorLevel.Error,
                //                        RowNum = bitDefRow.RowNum,
                //                        ColNum = bitDefRef.BitDefTable.LowLimitIdx + 1,
                //                        Message = $"Lo Limit mismatch {limitRow.TestName} : {lo} vs {limitRow.LowLimit}",
                //                    };
                //                    ErrorReportManager.AddError(error);
                //                }
                //            }
                //            if (bitDefRow.HighLimit != "N/A")
                //            {
                //                double hi = EfuseCmdUtility.ToDec(bitDefRow.HighLimit);
                //                if (hi != limitRow.HighLimit)
                //                {
                //                    var error = new Error()
                //                    {
                //                        ErrorType = BinCutErrorType.E_Formula,
                //                        SheetName = bitDefRow.SheetName,
                //                        ErrorLevel = ErrorLevel.Error,
                //                        RowNum = bitDefRow.RowNum,
                //                        ColNum = bitDefRef.BitDefTable.HighLimitIdx + 1,
                //                        Message = $"Hi Limit mismatch {limitRow.TestName} : {hi} vs {limitRow.HighLimit}",
                //                    };
                //                    ErrorReportManager.AddError(error);
                //                }
                //            }
                //        }

                //    }
                //}
                //#endregion

                colIdxMain++;
            }

            foreach (BitDefRow bitDefRow in bitDefRows)
            {
                if (!checkRows.Contains(bitDefRow) && !bitDefRow.Block.EqualsIgnoreCase("cmp_e") && !bitDefRow.Block.EqualsIgnoreCase("cmp_p"))
                {
                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_13, bitDefRow.SheetName, bitDefRow.RowNum, loaderEfuseBitDef.BitDefTable.Titles.Count + 1, [bitDefRow.Block, bitDefRow.Name]);
                    HighLightCell(excelWorksheet, bitDefRow.RowNum, loaderEfuseBitDef.BitDefTable.Titles.Count + 1, bitDefRow.RowNum, loaderEfuseBitDef.BitDefTable.Titles.Count + 1, Color.Red);
                }
            }
            #endregion
            return colIdxMain;
        }

        private static List<BitDefRow> PrintItem(LoaderEfuseBitDef loaderEfuseBitDef, ExcelWorksheet excelWorksheet)
        {
            #region Print item from EFUSE_BitDef_Table
            List<BitDefRow> bitDefRows = [];
            int rowIdx = 11;
            IEnumerable<IGrouping<string, List<string>>> groups = loaderEfuseBitDef.BitDefTable.Rows.GroupBy(x => x[loaderEfuseBitDef.BitDefTable.BlockIdx]);
            foreach (IGrouping<string, List<string>> group in groups)
            {
                // skip 1 row if items of block is different
                rowIdx++;
                string block = group.Key;
                for (int colIdx = 0; colIdx < loaderEfuseBitDef.BitDefTable.Titles.Count; colIdx++)
                {
                    excelWorksheet.Cells[rowIdx, colIdx + 1].Value = loaderEfuseBitDef.BitDefTable.Titles[colIdx];
                }
                excelWorksheet.Cells[rowIdx, 1].Value = block;
                rowIdx++;
                foreach (List<string> oneRow in group.ToList())
                {
                    PrintBdfContent(excelWorksheet, loaderEfuseBitDef.BitDefTable, oneRow, ref rowIdx);

                    if (!string.IsNullOrEmpty(oneRow[loaderEfuseBitDef.BitDefTable.NameIdx]))
                    {
                        BitDefRow bitDefRow = new BitDefRow { Block = block, Name = oneRow[loaderEfuseBitDef.BitDefTable.NameIdx] };
                        if (!bitDefRows.Any(x => x.Block.EqualsIgnoreCase(bitDefRow.Block) && x.Name.EqualsIgnoreCase(bitDefRow.Name)))
                        {
                            bitDefRow.Resolution = oneRow[loaderEfuseBitDef.BitDefTable.ResolutionIdx];
                            bitDefRow.RowNum = rowIdx;
                            bitDefRow.SheetName = excelWorksheet.Name;
                            bitDefRow.DefaultReal = oneRow[loaderEfuseBitDef.BitDefTable.DefaultOrRealIdx];
                            bitDefRow.IsDefault = oneRow[loaderEfuseBitDef.BitDefTable.DefaultOrRealIdx].EqualsIgnoreCase("Default");
                            bitDefRow.IsCrc = oneRow[loaderEfuseBitDef.BitDefTable.AlgorithmIdx].EqualsIgnoreCase("crc");
                            if (bitDefRow.IsCrc)
                            {
                                var crc = new CrcItem()
                                {
                                    FieldName = oneRow[loaderEfuseBitDef.BitDefTable.NameIdx],
                                    Blockindx = loaderEfuseBitDef.BitDefTable.BlockIdx,
                                    BlockName = oneRow[loaderEfuseBitDef.BitDefTable.BlockIdx],
                                    Description = oneRow[loaderEfuseBitDef.BitDefTable.DescriptionIdx],
                                    CrcMethod = oneRow[loaderEfuseBitDef.BitDefTable.BitWidthIdx],
                                    Job = oneRow[loaderEfuseBitDef.BitDefTable.PrgStageIdx]
                                };
                                crc.CalcBitwidth(int.Parse(oneRow[loaderEfuseBitDef.BitDefTable.MsbBitIdx]), int.Parse(oneRow[loaderEfuseBitDef.BitDefTable.LsbBitIdx]));
                                //UpdateCRCDescription(crc);
                                bitDefRow.Crc = crc;
                            }
                            bitDefRow.Type = bitDefRow.JudgeValueType(oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx]);
                            try
                            {
                                bitDefRow.DefaultValue = bitDefRow.IsDefault ? oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx] : "-999";
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(string.Format(ex.ToString()));
                            }
                            string lo = oneRow[loaderEfuseBitDef.BitDefTable.LowLimitIdx];
                            //EfuseCmdUtility.ConvertValue(lo);
                            bitDefRow.LowLimit = lo;
                            string hi = oneRow[loaderEfuseBitDef.BitDefTable.HighLimitIdx];
                            //EfuseCmdUtility.ConvertValue(hi);
                            bitDefRow.HighLimit = hi;
                            bitDefRow.RealStage = oneRow[loaderEfuseBitDef.BitDefTable.PrgStageIdx];
                            if (EfuseCmdUtility.JobStages.Any(x => bitDefRow.RealStage.Contains(x.Key)))
                            {
                                bitDefRow.JobStage = EfuseCmdUtility.JobStages.Where(x => bitDefRow.RealStage.Contains(x.Key)).Select(x => x.Value).First();
                            }
                            else
                            {
                                bitDefRow.JobStage = 999;
                            }

                            bitDefRow.Algorithm = oneRow[loaderEfuseBitDef.BitDefTable.AlgorithmIdx];
                            bitDefRow.Lsb = oneRow[loaderEfuseBitDef.BitDefTable.LsbBitIdx];
                            bitDefRow.Msb = oneRow[loaderEfuseBitDef.BitDefTable.MsbBitIdx];
                            bitDefRows.Add(bitDefRow);
                        }
                    }
                    rowIdx++;
                }
            }
            #endregion
            return bitDefRows;
        }

        private static string GetPrr(List<EfuseRow> efuseRows, string type)
        {
            bool isProber = false;
            if (type.ContainsIgnoreCase("prober"))
            {
                isProber = true;
            }

            EfuseRow? lotId = isProber ? efuseRows.Find(x => x.SubConfig.EqualsIgnoreCase("Prober_LotID")) : efuseRows.Find(x => x.SubConfig.ContainsIgnoreCase("LOT_ID") && x.BankConfig != null && x.BankConfig.ContainsIgnoreCase(type));
            EfuseRow? waferId = isProber ? efuseRows.Find(x => x.SubConfig.EqualsIgnoreCase("Prober_WaferID")) : efuseRows.Find(x => x.SubConfig.ContainsIgnoreCase("WAFER_ID") && x.BankConfig != null && x.BankConfig.ContainsIgnoreCase(type));
            EfuseRow? xCoord = isProber ? efuseRows.Find(x => x.SubConfig.EqualsIgnoreCase("Prober_X")) : efuseRows.Find(x => x.SubConfig.ContainsIgnoreCase("X_COORDINATE") && x.BankConfig != null && x.BankConfig.ContainsIgnoreCase(type));
            EfuseRow? yCoord = isProber ? efuseRows.Find(x => x.SubConfig.EqualsIgnoreCase("Prober_Y")) : efuseRows.Find(x => x.SubConfig.ContainsIgnoreCase("Y_COORDINATE") && x.BankConfig != null && x.BankConfig.ContainsIgnoreCase(type));

            if (lotId is null || waferId is null || xCoord is null || yCoord is null)
            {
                return string.Empty;
            }

            string combinedBits = lotId.Bits + waferId.Bits + xCoord.Bits + yCoord.Bits;
            string reversedBits = new([.. combinedBits.Reverse()]);

            return reversedBits;
        }

        public static void PrintHeader(LoaderEfuseBitDef loaderEfuseBitDef, ExcelWorksheet excelWorksheet)
        {
            excelWorksheet.Cells[1, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "SoftBin";
            excelWorksheet.Cells[2, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "HardBin";
            excelWorksheet.Cells[3, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "CurrentJob";
            excelWorksheet.Cells[4, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "XY_Handler";
            excelWorksheet.Cells[5, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "PRR_Code";
            excelWorksheet.Cells[6, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "EFuseLotNumber";
            excelWorksheet.Cells[7, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "EFuseWaferID";
            excelWorksheet.Cells[8, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "EFuseDieX";
            excelWorksheet.Cells[9, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "EFuseDieY";
            excelWorksheet.Cells[10, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "Hram_ECID_53bit";
            excelWorksheet.Cells[11, loaderEfuseBitDef.BitDefTable.Titles.Count].Value = "SVM_CFuse_288Bits";
        }

        private static string GetCrc(CrcItem crcItem, EfuseBitDefTable efuseBitDefTable, DiceInfo diceInfo)
        {
            List<List<string>> blockItems = [];
            List<List<string>> otherBankItems = [];
            string calcBank = "";
            if (crcItem.Description.Contains("BANK:", StringComparison.OrdinalIgnoreCase))
            {
                //calc other bank crc item
                calcBank = BankBracketRegex().Match(crcItem.Description).Groups["bank"].ToString();
                otherBankItems = efuseBitDefTable.Rows.FindAll(p => p[efuseBitDefTable.BlockIdx] == calcBank);
            }
            blockItems = efuseBitDefTable.Rows.FindAll(p => p[efuseBitDefTable.BlockIdx] == crcItem.BlockName);
            List<List<string>> sourceItems = string.IsNullOrEmpty(calcBank) ? blockItems : otherBankItems;
            int lsb = int.Parse(sourceItems.First()[efuseBitDefTable.LsbBitIdx]);
            int msb = int.Parse(sourceItems.First()[efuseBitDefTable.MsbBitIdx]);
            int offset = Math.Min(lsb, msb);
            List<string> keyList = string.IsNullOrEmpty(calcBank) ? [.. blockItems.Select(p => p[efuseBitDefTable.NameIdx])] : [.. otherBankItems.Select(p => p[efuseBitDefTable.NameIdx])];
            IEnumerable<string> ss = keyList.Where(x => x.Contains("crc", StringComparison.OrdinalIgnoreCase));
            (Dictionary<string, string> bitStreams, List<string> wrongBitWidthItems, Dictionary<string, string> missingitems) = GetBits(crcItem, efuseBitDefTable, diceInfo, calcBank, keyList);

            string result = "";
            if (bitStreams.Count != 0 && crcItem.BlockName != "UID")
            {
                crcItem.RawBitstream = string.Join("", bitStreams.Values.Reverse());
                crcItem.Offset = offset;
                if (crcItem.Description.StartsWithIgnoreCase("ONE_COMPLEMENT_TARGET"))
                {
                    string refFieldName = FieldNameBracketRegex().Match(crcItem.Description).Groups["fieldName"].ToString();
                    diceInfo.Items.TryGetValue("Bank_" + crcItem.BlockName + "#" + refFieldName, out EfuseDatalogItem? datalogItem);
                    List<string>? bdfItem = efuseBitDefTable.Rows.FirstOrDefault(p => p[efuseBitDefTable.NameIdx].EqualsIgnoreCase(refFieldName) && p[efuseBitDefTable.BlockIdx].EqualsIgnoreCase(crcItem.BlockName));
                    if (datalogItem != null && bdfItem != null)
                    {
                        result = CrcItem.CalcCrcWithOneComplement(datalogItem.RawData, int.Parse(bdfItem[efuseBitDefTable.BitWidthIdx]));
                    }
                }
                else
                {
                    result = crcItem.CalcCrc();
                }
            }
            return result;
        }

        private static (Dictionary<string, string> bitStream, List<string> missingitems, Dictionary<string, string> wrong_bit_width_items) GetBits(CrcItem crcItem, EfuseBitDefTable efuseBitDefTable, DiceInfo diceInfo, string calcBank, List<string> keyList)
        {
            Dictionary<string, string> bitStream = [];
            List<string> missingitems = [];
            Dictionary<string, string> wrongBitWidthItems = [];
            foreach (string key in keyList)
            {
                string bank = string.IsNullOrEmpty(calcBank) ? "Bank_" + crcItem.BlockName : "Bank_" + calcBank;
                diceInfo.Items.TryGetValue(bank + "#" + key, out EfuseDatalogItem? datalogItem);
                if (datalogItem != null)
                {
                    if (!string.IsNullOrEmpty(datalogItem.RawData))
                    {
                        bitStream.Add(key, datalogItem.RawData);
                    }
                    List<string> bdfItem = efuseBitDefTable.Rows.FirstOrDefault(p => ("Bank_" + p[efuseBitDefTable.BlockIdx]).EqualsIgnoreCase(datalogItem.Block) && p[efuseBitDefTable.NameIdx].EqualsIgnoreCase(datalogItem.Id))!;
                    if (int.Parse(bdfItem[efuseBitDefTable.BitWidthIdx]) != datalogItem.RawData.Length)
                    {
                        wrongBitWidthItems.Add(datalogItem.Id, datalogItem.RawData);
                        string message = $"The bit width of {datalogItem.Id} : {string.Join("", datalogItem.RawData)} is not match table {bdfItem[efuseBitDefTable.BitWidthIdx]} !!!";
                        Console.WriteLine(message);
                        ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_14, "", 0, 0, [datalogItem.Id, string.Join("", datalogItem.RawData), bdfItem[efuseBitDefTable.BitWidthIdx]]);
                    }
                }
                else
                {
                    missingitems.Add(key);
                    string message = $"Can not find {key} in log !!!";
                    Console.WriteLine(message);
                    ErrorReportManager.AddError(EfuseCheckCmdLibError.E_MismatchValue_15, "", 0, 0, [key]);
                }
            }
            return (bitStream, missingitems, wrongBitWidthItems);
        }

        public static void PrintCfgTableHeader(ExcelWorksheet excelWorksheet, EfuseCfgTable efuseCfgTable, ref int colIndx, string scenario)
        {
            int rowindx = 4;
            colIndx = 1;
            if (!efuseCfgTable.Scenario.TryGetValue(scenario, out int cfgScenarioIndx))
            {
                Console.WriteLine("Config didn't find! " + scenario);
                return;
            }

            #region
            var indexList = new List<int> { efuseCfgTable.ConditionIdx, efuseCfgTable.MsbBitIdx, efuseCfgTable.LsbBitIdx, efuseCfgTable.BitWidthIdx, efuseCfgTable.PrgStageIdx, cfgScenarioIndx };
            #endregion
            //print Scenario(CFG and Datalog)
            excelWorksheet.Cells[1, colIndx].Value = $"Scenario in ConfigTable:{scenario}, in Datalog:{scenario}";
            excelWorksheet.Cells[2, colIndx].Value = "Flag DisableChkLMT:False";
            //print CFG Table Header
            foreach (int item in indexList)
            {
                excelWorksheet.Cells[rowindx, colIndx].Value = efuseCfgTable.TitleList[item];
                colIndx++;
            }
            rowindx++;

            //print CFG Table Content
            foreach (List<string> rows in efuseCfgTable.CfgRows)
            {
                colIndx = 1;
                if (rows.Count < cfgScenarioIndx)
                {
                    continue;
                }

                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.ConditionIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.MsbBitIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.LsbBitIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.BitWidthIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[efuseCfgTable.PrgStageIdx];
                colIndx++;
                excelWorksheet.Cells[rowindx, colIndx].Value = rows[cfgScenarioIndx];

                rowindx++;
            }
        }

        public static void PrintBdfContent(ExcelWorksheet excelWorksheet, EfuseBitDefTable efuseBitDefTable, List<string> oneRow, ref int lastAdRow)
        {
            for (int colIdx = 0; colIdx < efuseBitDefTable.Titles.Count; colIdx++)
            {
                if (colIdx < oneRow.Count)
                {
                    if (FormulaPrefixRegex().IsMatch(oneRow[colIdx]))
                    {
                        excelWorksheet.Cells[lastAdRow, colIdx + 1].Formula = oneRow[colIdx];
                    }
                    else
                    {
                        excelWorksheet.Cells[lastAdRow, colIdx + 1].Value = oneRow[colIdx];
                    }
                }
            }

            if (oneRow[efuseBitDefTable.DefaultOrRealIdx].EqualsIgnoreCase("real"))
            {
                HighLightCell(excelWorksheet, lastAdRow, efuseBitDefTable.DefaultOrRealIdx + 1, lastAdRow, efuseBitDefTable.DefaultOrRealIdx + 1, Color.PaleGoldenrod);
            }
        }

        public static void HighLightCell(ExcelWorksheet excelWorksheet, int stRow, int stCol, int spRow, int spCol, Color color)
        {
            ExcelRange rng = excelWorksheet.SelectedRange[stRow, stCol, spRow, spCol];
            rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(color);
        }
    }
}
