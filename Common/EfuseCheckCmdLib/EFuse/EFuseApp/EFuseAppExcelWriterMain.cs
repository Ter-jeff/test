using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Static;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using EfuseCheckCmdLib.AlgorithmCheck;
using EfuseCheckCmdLib.Static;
using EfuseCheckCmdLib.Utility;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public partial class EFuseAppExcelWriterMain(EfuseScriptConfig efuseScriptConfig) : EFuseAppExcelWriter(efuseScriptConfig)
    {

        private const string FusePattenSummarySheetName = "FusePattenSummary";
        [GeneratedRegex("Default", RegexOptions.IgnoreCase)]
        private static partial Regex DefaultRegex();

        [GeneratedRegex("base", RegexOptions.IgnoreCase)]
        private static partial Regex BaseRegex();

        [GeneratedRegex("CP", RegexOptions.IgnoreCase)]
        private static partial Regex CpRegex();

        [GeneratedRegex("ids", RegexOptions.IgnoreCase)]
        private static partial Regex IdsRegex();

        [GeneratedRegex(@"\d\s*ma", RegexOptions.IgnoreCase)]
        private static partial Regex MilliampsRegex();

        [GeneratedRegex("PRODUCT_IDENTIFIER|^IDS_VDD", RegexOptions.IgnoreCase)]
        private static partial Regex ProductIdentifierOrIdsVddRegex();

        public void WriteMainFile(Action<string, string> appendRichText, string filename, LoaderEfuseBitDef loaderEfuseBitDef, List<XParseDatalog.StageDicesInfo> stageDicesInfos, List<XParseDatalog.EFuseSyntaxChkItem> eFuseSyntaxChkItems)
        {
            EfuseCheckResultType result = EfuseCheckResultType.Pass;
            bool isExistDsscRead = stageDicesInfos.SelectMany(p => p.AlldicesDiceInfos).ToList().Any(p => p.AllReadFromDssc.Count > 0);
            if (!isExistDsscRead)
            {
                return;
            }

            EFuseAppExcelWriterStatic.SummaryErrorsList.Clear();
            bool isExistDssc = stageDicesInfos.SelectMany(p => p.AlldicesDiceInfos).ToList().Any(p => p.AllReadFromDssc.Count > 0);
            if (!isExistDssc)
            {
                return;
            }

            FileInfo newFile = Writer.CreateNewFile(filename);
            WriteReport(appendRichText, newFile, loaderEfuseBitDef, stageDicesInfos, eFuseSyntaxChkItems);

            if (EFuseAppExcelWriterStatic.SummaryErrorsList.Count != 0)
            {
                result = EfuseCheckResultType.Fail;
            }

            if (EfuseStatic.Result == EfuseCheckResultType.Pass)
            {
                EfuseStatic.Result = result;
            }
        }

        protected override void PrintOthers(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, List<XParseDatalog.EFuseSyntaxChkItem> eFuseSyntaxChkItems, ExcelWorksheet excelWorksheet, ref int lastAdRow, ref string curBlock)
        {
            foreach (List<string> oneRow in loaderEfuseBitDef.BitDefTable.Rows)
            {
                if (curBlock != oneRow[loaderEfuseBitDef.BitDefTable.BlockIdx])
                {
                    // skip 1 row if items of block is different
                    lastAdRow++;
                    Writer.PrintBlockHeader(excelWorksheet, loaderEfuseBitDef.BitDefTable, oneRow, ref curBlock, ref lastAdRow);
                }
                var oneBitRow = new XBitDefRow { Block = oneRow[loaderEfuseBitDef.BitDefTable.BlockIdx] };
                oneBitRow.Name = oneBitRow.Block + "#" + oneRow[loaderEfuseBitDef.BitDefTable.NameIdx];
                oneBitRow.Resolution = oneRow[loaderEfuseBitDef.BitDefTable.ResolutionIdx];
                oneBitRow.RowNo = lastAdRow;
                oneBitRow.DefaultReal = oneRow[loaderEfuseBitDef.BitDefTable.DefaultOrRealIdx];
                oneBitRow.IsDefault = DefaultRegex().IsMatch(oneRow[loaderEfuseBitDef.BitDefTable.DefaultOrRealIdx]);
                try
                {
                    oneBitRow.DefaultValue = oneBitRow.IsDefault
                        ? oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx]
                        : "-999";
                }
                catch (Exception ex)
                {
                    EfuseStatic.Result = EfuseCheckResultType.Exception;
                    if (EfuseStatic.IsCmd)
                    {
                        appendRichText(string.Format(ex.ToString()), "Red");
                    }
                    else
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                }
                oneBitRow.Type = oneBitRow.IsDefault
                    ? EfuseScriptUtility.JudgeValueType(oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx])
                    : EnumValueType.Other;
                if (!_config.UdrcmpMapping.ContainsKey(oneBitRow.Block) && oneBitRow.Block.Contains("CMP"))
                {
                    _config.UdrcmpMapping.Add(oneBitRow.Block, oneBitRow.Block.Replace("CMP", "UDR"));
                }
                if (XParseDatalog.StdfReader != null)
                {
                    #region compare stdf HIP
                    if (loaderEfuseBitDef.BitDefTable.HipNameIdx != -1)
                    {
                        string hipName = oneRow[loaderEfuseBitDef.BitDefTable.HipNameIdx];
                        if (!string.IsNullOrEmpty(hipName) && XParseDatalog.StdfReader.HighLDic.TryGetValue(hipName, out string? value))
                        {
                            oneBitRow.HiPparName = hipName;
                            oneBitRow.HiPparEq = oneRow[loaderEfuseBitDef.BitDefTable.HipEquationIdx];
                            oneRow.Add(value);
                            oneRow.Add(XParseDatalog.StdfReader.LowLDic[hipName]);
                        }
                        else
                        {
                            oneRow.Add("");
                            oneRow.Add("");
                        }
                        oneBitRow.HipHighLimit =
                            EfuseScriptUtility.ConvertValue(appendRichText, oneRow[loaderEfuseBitDef.BitDefTable.HiplimithIdx]);
                        oneBitRow.HipLowLimit =
                            EfuseScriptUtility.ConvertValue(appendRichText, oneRow[loaderEfuseBitDef.BitDefTable.HiplimitlIdx]);
                    }
                    #endregion
                }
                PrintBdfContent(eFuseSyntaxChkItems, testCat, appendRichText, excelWorksheet, loaderEfuseBitDef.BitDefTable, oneRow, ref lastAdRow, _config);
                #region Check Base Voltage with Config
                if (BaseRegex().IsMatch(oneRow[loaderEfuseBitDef.BitDefTable.AlgorithmIdx]))
                {
                    if (oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx].ContainsIgnoreCase("b"))
                    {
                        EfuseScriptUtility.BaseVoltage = Convert.ToInt16(oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx].ToLower().Replace("b0", "").Replace("b", "").Replace("_", ""), 2);
                    }
                    else
                    {
                        EfuseScriptUtility.BaseVoltage = int.Parse(oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx]);
                    }

                    EfuseScriptUtility.BaseVoltageResolution =
                        double.TryParse(oneRow[loaderEfuseBitDef.BitDefTable.ResolutionIdx], out double resolution)
                            ? resolution
                            : _config.BaseVoltageResolution;

                    if (EfuseScriptUtility.BaseVoltage != _config.BaseVoltage)
                    {
                        var error = new Error(EfuseCheckCmdLibError.W_MismatchValue_01, EnumErrorLevel.Warning, "", lastAdRow, loaderEfuseBitDef.BitDefTable.DefaultValueIdx + 1, "Not Match to Config BaseVoltage " + _config.BaseVoltage);
                        Writer.HighLightCell(excelWorksheet, lastAdRow, loaderEfuseBitDef.BitDefTable.DefaultValueIdx + 1, lastAdRow, loaderEfuseBitDef.BitDefTable.DefaultValueIdx + 1, Color.Yellow);
                        AddComment(appendRichText, excelWorksheet, error);
                    }
                }
                #endregion

                if (!string.IsNullOrEmpty(oneBitRow.Name.Split('#')[1]))
                {
                    if (!XReadEfuseBitDef.Contains(oneBitRow.Name))
                    {
                        string lLimStr = oneRow[loaderEfuseBitDef.BitDefTable.LowLimitIdx];
                        oneBitRow.LowLimit = EfuseScriptUtility.ConvertValue(appendRichText, lLimStr);
                        string hLimStr = oneRow[loaderEfuseBitDef.BitDefTable.HighLimitIdx];
                        oneBitRow.HighLimit = EfuseScriptUtility.ConvertValue(appendRichText, hLimStr);
                        oneBitRow.RealStage = oneRow[loaderEfuseBitDef.BitDefTable.PrgStageIdx];
                        oneBitRow.JobStage = EfuseScriptUtility.GetJobOrder(oneBitRow.RealStage);
                        oneBitRow.Algorithm = oneRow[loaderEfuseBitDef.BitDefTable.AlgorithmIdx];
                        oneBitRow.Lsb = oneRow[loaderEfuseBitDef.BitDefTable.LsbBitIdx];
                        oneBitRow.Msb = oneRow[loaderEfuseBitDef.BitDefTable.MsbBitIdx];
                        //
                        XReadEfuseBitDef.BitDefTable.Add(oneBitRow);
                    }
                }

                lastAdRow++;
            }
        }

        protected override void PrintSetWrite(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, List<XParseDatalog.StageDicesInfo> stageDicesInfos, ref int testCat, Dictionary<string, List<double>> limitLimitation, ref List<string> bdfRealItems, List<List<CrcItem>> totalCrcItems, ref string runStage, bool isContainsHipInfo, bool isContainsBvInfo, List<XBitDefRow> realValueTable, List<XBitDefRow> setWriteValueTable, ExcelPackage excelPackage, ExcelWorksheet summaryWorkSheet, ref ExcelWorksheet bdfWorksheet, ExcelWorksheet hardipWorkSheet, ExcelWorksheet setWriteVariableSheet, ExcelWorksheet eccCheckSheet)
        {
            int colIdxMain = loaderEfuseBitDef.BitDefTable.Titles.Count + 1;
            int colIdxMainNext = colIdxMain;
            int colIdxConfig = 1;
            int colidxRealValue = 5;
            int colidxSetWrite = 1;
            int eccDataStartCol = 2;
            int count = 0;
            var checksumList = new List<CheckSum>();

            foreach (XParseDatalog.StageDicesInfo allDiceInfo in stageDicesInfos)
            {
                bdfWorksheet = excelPackage.Workbook.Worksheets[SheetConst.Type5BitDefTable];
                testCat = EfuseScriptUtility.GetJobOrder(allDiceInfo.Stage);
                runStage = allDiceInfo.Stage;
                long idie = 0;
                int countForDie = 0;
                count++;
                bool isTheLast = count == stageDicesInfos.Count;

                foreach (XDiceInfo oneDice in allDiceInfo.AlldicesDiceInfos)
                {
                    if (oneDice.ChecksumList.Count != 0)
                    {
                        checksumList = oneDice.ChecksumList;
                    }
                    var deviceRealItems = new List<string>();
                    idie++;
                    try
                    {
                        countForDie++;
                        bool isTheLastForDie = countForDie == allDiceInfo.AlldicesDiceInfos.Count;
                        #region HIP Check
                        appendRichText(oneDice.XCoor + "," + oneDice.YCoor + " => " + (int)(idie * 100 / allDiceInfo.AlldicesDiceInfos.Count) + "%", "Blue");
                        HipItem? stdfDice = null;
                        if (XParseDatalog.StdfReader != null)
                        {
                            stdfDice = XParseDatalog.StdfReader.HipItems.LastOrDefault(p => p.X.Trim() == oneDice.XCoor.ToString() && p.Y.Trim() == oneDice.YCoor.ToString());
                        }
                        #endregion

                        oneDice.AllReadFromDssc.AddRange(oneDice.AllUdrVer);
                        if (!EfuseScriptUtility.IsPassBin(oneDice.SortBin, _config) && EfuseStatic.ShowType == 1)
                        {
                            continue;
                        }

                        if (EfuseScriptUtility.IsPassBin(oneDice.SortBin, _config) && EfuseStatic.ShowType == 2)
                        {
                            continue;
                        }

                        if (stdfDice != null)
                        {
                            colIdxMainNext += 3;
                        }
                        else
                        {
                            colIdxMainNext++;
                        }
                        PrintHeader(bdfWorksheet, colIdxMain, allDiceInfo, oneDice);

                        CompareXy(appendRichText, oneDice, runStage, bdfWorksheet, colIdxMain);
                        var itemList = XReadEfuseBitDef.BitDefTable.Where(p => !p.Algorithm.EqualsIgnoreCase("cond")).Select(p => p.RowNo).ToList();

                        PrintData(appendRichText, loaderEfuseBitDef, testCat, limitLimitation, ref bdfRealItems, totalCrcItems, isContainsHipInfo, isContainsBvInfo, realValueTable, setWriteValueTable, excelPackage, bdfWorksheet, hardipWorkSheet, setWriteVariableSheet, eccCheckSheet, colIdxMain, ref colidxRealValue, ref colidxSetWrite, ref eccDataStartCol, checksumList, allDiceInfo, countForDie, isTheLast, oneDice, ref deviceRealItems, isTheLastForDie, stdfDice, itemList);
                    }
                    catch (Exception e)
                    {
                        EfuseStatic.Result = EfuseCheckResultType.Exception;
                        if (EfuseStatic.IsCmd)
                        {
                            appendRichText(string.Format(e.ToString()), "Red");
                        }
                        else
                        {
                            ErrorMessageBox.Show(string.Format(e.ToString()));
                        }
                    }
                    foreach (Error generalError in EFuseAppExcelWriterStatic.GeneralCheckerErrors)
                    {
                        if (!string.IsNullOrEmpty(generalError.Message))
                        {
                            AddComment(appendRichText, bdfWorksheet, generalError);
                        }
                    }
                    colIdxMain = colIdxMainNext;
                }
                //Create Harvest sheet field sheet
                Writer.SetHarvestSheetSummary(excelPackage, bdfWorksheet);
                SetSheetSummaryReport(summaryWorkSheet);

                // If user import Config Table, tool would generate the Config Table Sheet and check with datalog
                // Generally, this flag should be true...
                if (EfuseCfgTableReader1.IsContainCfgTable)
                {
                    bdfWorksheet = CheckCFGTable(appendRichText, loaderEfuseBitDef, excelPackage, ref colIdxConfig, allDiceInfo);
                }
            }
        }

        private void PrintData(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, Dictionary<string, List<double>> limitLimitation, ref List<string> bdfRealItems, List<List<CrcItem>> totalCrcItems, bool isContainsHipInfo, bool isContainsBvInfo, List<XBitDefRow> realValueTable, List<XBitDefRow> setWriteValueTable, ExcelPackage excelPackage, ExcelWorksheet bdfWorksheet, ExcelWorksheet hardipWorkSheet, ExcelWorksheet setWriteVariableSheet, ExcelWorksheet eccCheckSheet, int colIdxMain, ref int colidxRealValue, ref int colidxSetWrite, ref int eccDataStartCol, List<CheckSum> checkSums, XParseDatalog.StageDicesInfo stageDicesInfo, int countForDie, bool isTheLast, XDiceInfo xDiceInfo, ref List<string> deviceRealItems, bool isTheLastForDie, HipItem? hipItem, List<int> itemList)
        {
            #region print data
            EFuseAppExcelWriterStatic.GeneralCheckerErrors.Clear();
            bool isNeedIdsCheck = xDiceInfo.IdsFuseInfo.Count > 0;
            //Warning: It's not align calculation with VBT, so tool need reverse for this case
            bool isNeedReversedPrr = true;
            bool isNeedReversed = true;
            var dsscInfo = new DeviceInfo();
            var lotIdList = new List<string>();

            LoopAllReadFromDssc(appendRichText, loaderEfuseBitDef, testCat, bdfWorksheet, colIdxMain, xDiceInfo, hipItem, itemList, ref isNeedReversedPrr, ref isNeedReversed, dsscInfo, ref lotIdList);

            xDiceInfo.PrrLotId = string.Join("", lotIdList);
            dsscInfo.Lot = string.Join("", lotIdList);
            if ((xDiceInfo.Prober.XCor != dsscInfo.XCor || xDiceInfo.Prober.YCor != dsscInfo.YCor || xDiceInfo.Prober.Lot != dsscInfo.Lot || xDiceInfo.Prober.Wafer != dsscInfo.Wafer) && CpRegex().IsMatch(stageDicesInfo.Stage))
            {
                var xyError = new Error(EfuseCheckCmdLibError.E_MismatchValue_27, EnumErrorLevel.Error, "", 4, colIdxMain, $"Prober:{xDiceInfo.Prober.Lot}_{xDiceInfo.Prober.Wafer}_{xDiceInfo.Prober.XCor}_{xDiceInfo.Prober.YCor};DSSC:{dsscInfo.Lot}_{dsscInfo.Wafer}_{dsscInfo.XCor}_{dsscInfo.YCor} not same");
                xyError.Comments.Add(xyError.Message);
                EFuseAppExcelWriterStatic.GeneralCheckerErrors.Add(xyError);
            }

            if (EfuseCfgTableReader1.IsContainCfgTable)
            {
                EfuseCfgTable table = EfuseScriptUtility.CfgTableSel();
                CheckSvmcFuse(appendRichText, bdfWorksheet, xDiceInfo, stageDicesInfo, colIdxMain, table);
            }

            if (isContainsHipInfo || isContainsBvInfo || isNeedIdsCheck)
            {
                CheckAll(appendRichText, bdfRealItems, isContainsHipInfo, isContainsBvInfo, realValueTable, hardipWorkSheet, ref colidxRealValue, xDiceInfo, ref deviceRealItems, isNeedIdsCheck);
            }

            if (setWriteVariableSheet != null && xDiceInfo.HardIpFuseInfo.Count > 0 && xDiceInfo.HardIpFuseInfo.Select(p => p.Value.ReferenceValue).Any())
            {
                WriteSetWriteValueSheet(appendRichText, setWriteVariableSheet, setWriteValueTable, xDiceInfo, ref colidxSetWrite, limitLimitation, ref bdfRealItems);
            }
            if (isTheLastForDie && xDiceInfo.FusePatternLines.Count != 0)
            {
                ExcelWorksheet fusePatternSummarySheet = excelPackage.Workbook.Worksheets[FusePattenSummarySheetName] ??
                                                         excelPackage.Workbook.Worksheets.Add(FusePattenSummarySheetName);
                Writer.WriteFusePatternSummarySheet(appendRichText, fusePatternSummarySheet, xDiceInfo);
            }
            if (eccCheckSheet != null && xDiceInfo.EccInfo.Count > 0 && xDiceInfo.EccInfo.Select(p => p.Value.RawDataList).Any())
            {
                WriteEccCheckSheet(appendRichText, eccCheckSheet, xDiceInfo, loaderEfuseBitDef.BitDefTable, ref eccDataStartCol);
            }
            if (isTheLast && isTheLastForDie)
            {
                if (countForDie > 1)
                {   //Need check the case if the site count is > 1
                    CheckAllSiteRealValue(appendRichText, loaderEfuseBitDef, stageDicesInfo.Stage, bdfWorksheet, hardipWorkSheet, setWriteVariableSheet);
                }
            }

            CheckNotFoundInDatalog(appendRichText, bdfWorksheet, colIdxMain, itemList);

            //Check efuse rule : CRCCheck
            List<CrcItem> crcInfo = CrcCheck(appendRichText, loaderEfuseBitDef, testCat, bdfWorksheet, colIdxMain, xDiceInfo);
            totalCrcItems.Add(crcInfo);

            foreach (CheckSum checksum in checkSums)
            {
                EfuseScriptUtility.CalcCheckSum(appendRichText, bdfWorksheet, checksum, xDiceInfo, loaderEfuseBitDef.BitDefTable, colIdxMain);
            }
            CheckPrr(appendRichText, bdfWorksheet, xDiceInfo, colIdxMain, isNeedReversedPrr);
            CheckIedaNull(appendRichText, bdfWorksheet, colIdxMain);
            //PRR Code Calculate?
            CheckHramEcid(appendRichText, bdfWorksheet, xDiceInfo, colIdxMain);
            CheckSvmcFuseSiteFail(appendRichText, bdfWorksheet, xDiceInfo, colIdxMain);
            #endregion
        }

        private static void CheckNotFoundInDatalog(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int colIdxMain, List<int> itemList)
        {
            foreach (int rowNo in itemList)
            {
                var missingError = new Error(EfuseCheckCmdLibError.E_MissingValue_01, EnumErrorLevel.Error, "", rowNo, colIdxMain, EfuseCheckCmdLibError.E_MissingValue_01.MessageTemplate);
                PrintErrorComment(appendRichText, excelWorksheet, missingError);
            }
        }

        private void LoopAllReadFromDssc(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, ExcelWorksheet excelWorksheet, int colIdxMain, XDiceInfo xDiceInfo, HipItem? hipItem, List<int> itemList, ref bool isNeedReversedPrr, ref bool isNeedReversed, DeviceInfo deviceInfo, ref List<string> lotIdList)
        {
            foreach (EfuseDatalogItem onePair in xDiceInfo.AllReadFromDssc)
            {
                XBitDefRow? row = XReadEfuseBitDef.GetRow(onePair.Id, onePair.Block);
                //isNeedReversed = (int.Parse(row.LSB) > int.Parse(row.MSB));
                isNeedReversed = int.TryParse(row?.Lsb, out int lsb) && int.TryParse(row?.Msb, out int msb) && lsb > msb;

                PreSetup(xDiceInfo, ref isNeedReversedPrr, deviceInfo, ref lotIdList, onePair, row!);

                string blockname = EfuseScriptUtility.CheckAliasBlock(onePair.Block, _config);
                //if (bitDefMain.Contains(idName)) => used to mapping item with Category and Block
                if (row != null)
                {
                    itemList.Remove(row.RowNo);
                    excelWorksheet.Cells[row.RowNo, colIdxMain].Value = row.IsDefault ? EfuseScriptUtility.ConvertValueToSpecifyFormat(appendRichText, onePair.Value, XReadEfuseBitDef.GetType(onePair.Id, blockname)) : onePair.Value;
                    if (hipItem != null)
                    {
                        CheckHip(appendRichText, excelWorksheet, colIdxMain, hipItem, onePair, row, blockname);
                    }
                    //if id equals to config condition, need to check with scenario
                    if (EfuseCfgTableReader1.IsContainCfgTable)
                    {
                        var tableItems = EfuseScriptUtility.CfgTableSel().CfgRows.Select(p => p[0]).ToList();
                        if (tableItems.Exists(p => p.EqualsIgnoreCase(row.Name.Split('#')[1])))
                        {
                            continue;
                        }
                    }
                    //Check efuse rule1: Check H/L limit and check real value equal to default if property of "Default or Real" is Default
                    if (row.JobStage <= testCat)
                    {
                        SetByJobStage1(appendRichText, loaderEfuseBitDef, testCat, excelWorksheet, xDiceInfo, isNeedReversed, onePair, row, blockname, row, colIdxMain);
                    }
                    else // add to check value is 0 if the items exceed to current stage
                    {
                        CheckIfAfterProgrammingStage(appendRichText, onePair, row, colIdxMain, excelWorksheet);
                    }
                }
            }
        }

        private static void CheckAll(Action<string, string> appendRichText, List<string> bdfRealItems, bool isContainsHipInfo, bool isContainsBvInfo, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, ref int colidxRealValue, XDiceInfo xDiceInfo, ref List<string> deviceRealItems, bool isNeedIdsCheck)
        {
            deviceRealItems.AddRange(bdfRealItems);
            deviceRealItems = [.. deviceRealItems.Distinct()];
            excelWorksheet.Cells[1, colidxRealValue, 1, colidxRealValue + 1].Value =
                $"X{xDiceInfo.XCoor}_Y{xDiceInfo.YCoor}";
            excelWorksheet.Cells[2, colidxRealValue].Value = "Real";
            excelWorksheet.Cells[2, colidxRealValue + 1].Value = "eFuse";
            //BDF HIP dsscLog
            int rowoffset = 3;
            if (isContainsHipInfo)
            {
                CheckByHipInfo2(appendRichText, xBitDefRows, excelWorksheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset);
            }
            if (isNeedIdsCheck)
            {
                CheckIds(appendRichText, xBitDefRows, excelWorksheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset);
            }

            if (isContainsBvInfo)
            {
                CheckBv(appendRichText, xBitDefRows, excelWorksheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset);
            }
            foreach (string deviceRealItem in deviceRealItems)
            {
                int rowindexReal = xBitDefRows.FindIndex(p => p.Name.EqualsIgnoreCase(deviceRealItem));
                Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);

                var error = new Error(EfuseCheckCmdLibError.E_MissingValue_01, EnumErrorLevel.Error, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, EfuseCheckCmdLibError.E_MissingValue_01.MessageTemplate);
                AddComment(appendRichText, excelWorksheet, error);
            }

            colidxRealValue += 2;
        }

        private static ExcelWorksheet CheckCFGTable(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, ExcelPackage excelPackage, ref int colIdxConfig, XParseDatalog.StageDicesInfo stageDicesInfo)
        {
            ExcelWorksheet bdfWorksheet;
            //Check CFGTable

            EfuseCfgTable table = EfuseScriptUtility.CfgTableSel();
            if (excelPackage.Workbook.Worksheets[table.TableName] == null)
            {
                bdfWorksheet = excelPackage.Workbook.Worksheets.Add(table.TableName);
                Writer.PrintCfgTableHeader(appendRichText, bdfWorksheet, table, ref colIdxConfig);
                colIdxConfig++;
            }
            bdfWorksheet = excelPackage.Workbook.Worksheets[table.TableName];
            XParseDatalog.TestCat = EfuseScriptUtility.GetJobOrder(stageDicesInfo.Stage);

            foreach (XDiceInfo oneDice in stageDicesInfo.AlldicesDiceInfos)
            {
                EfuseScriptUtility.CheckCfgCondition(appendRichText, bdfWorksheet, oneDice, loaderEfuseBitDef.BitDefTable, table, colIdxConfig);
                colIdxConfig++;
            }

            return bdfWorksheet;
        }

        private static void CheckByHipInfo2(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, List<string> deviceRealItems, int rowoffset)
        {
            var hipDsscItems = xDiceInfo.AllReadFromDssc.Where(p => xDiceInfo.HardIpFuseInfo.Keys.Any(q => p.Id.EqualsIgnoreCase(q.ToLower().Split('#')[1]) && p.Block.EqualsIgnoreCase(q.ToLower().Split('#')[0]))).ToList();

            foreach (EfuseDatalogItem hipDsscItem in hipDsscItems)
            {
                deviceRealItems.Remove(hipDsscItem.Block + "#" + hipDsscItem.Id.ToUpper());
                try
                {
                    int rowindexReal =
                        xBitDefRows.FindIndex(
                            p =>
                                p.Name.EqualsIgnoreCase(hipDsscItem.Block + "#" + hipDsscItem.Id));

                    string hipTarget =
                        xDiceInfo.HardIpFuseInfo.Keys.FirstOrDefault(
                            p =>
                                p.EqualsIgnoreCase(hipDsscItem.Block + "#" + hipDsscItem.Id))!;
                    if (rowindexReal != -1)
                    {
                        excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue].Value =
                            xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.TrimStart().Split(' ')[0];
                        excelWorksheet.Cells[
                            rowindexReal + rowoffset, colidxRealValue + 1]
                            .Value
                            = hipDsscItem.Value;
                        if (xBitDefRows[rowindexReal].Algorithm.Trim().EqualsIgnoreCase("ids"))
                        {
                            string regIds = @"\(\s*(?<value>[\-\w\.]+)";
                            string idsValue = Regex.Match(xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue, regIds, RegexOptions.IgnoreCase).Groups["value"].Value;

                            double resolution =
                                double.Parse(xBitDefRows[rowindexReal].Resolution);
                            double idsMeasValue = double.Parse(idsValue);
                            double expectIdsRealValue = EfuseStatic.IdsRoundMethod
                                ? Math.Floor(idsMeasValue / resolution) * resolution
                                : Math.Ceiling(idsMeasValue / resolution) * resolution;
                            double tempValue = double.Parse(idsValue) / resolution;
                            if (Math.Abs(expectIdsRealValue - double.Parse(hipDsscItem.Value)) > 0.01 &&
                                Math.Abs(tempValue - double.Parse(xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.TrimStart().Split(' ')[0])) > 0.01)
                            {
                                var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, null!);
                                Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);
                                AddComment(appendRichText, excelWorksheet, error);
                            }

                        }
                        else
                        {

                            if (!xBitDefRows[rowindexReal].Algorithm.Trim().EqualsIgnoreCase("vddbin"))
                            {
                                string value = xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.Split(' ')[0];

                                value = value.IsHexValue()
                                    ? Convert.ToInt64(value, 16).ToString()
                                    : value;
                                string hipDsscValue = hipDsscItem.Value.IsHexValue()
                                        ? Convert.ToInt64(hipDsscItem.Value, 16).ToString()
                                        : hipDsscItem.Value;
                                if (Math.Round(double.Parse(value)) != Math.Round(double.Parse(hipDsscValue)))
                                {
                                    var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, null!);
                                    AddComment(appendRichText, excelWorksheet, error);

                                    Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);

                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    EfuseStatic.Result = EfuseCheckResultType.Exception;
                    if (EfuseStatic.IsCmd)
                    {
                        appendRichText(string.Format(ex.ToString()), "Red");
                    }
                    else
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                }

            }
        }

        private void SetByJobStage1(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, bool isNeedReversed, EfuseDatalogItem efuseDatalogItem, XBitDefRow xBitDefRow, string blockname, XBitDefRow row, int colIdxMain)
        {
            if (_config.UdrcmpMapping.ContainsKey(blockname))
            {
                #region UDR CMP

                if (!EfuseScriptUtility.CompareUdrcmpItems(appendRichText, blockname, efuseDatalogItem.Id, xDiceInfo, loaderEfuseBitDef, _config))
                {
                    var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, EnumErrorLevel.Error, "", row.RowNo, colIdxMain, "Compared item not match.");
                    PrintErrorComment(appendRichText, excelWorksheet, error);
                    //PrintErrorComment(worksheet, Check_error, Color.Red);
                }
                #endregion
            }
            else
            {
                #region non UDR CMP items
                bool isFail = false;
                string err = "Fail: ";
                //if (onePair.RawData.Length>64)
                //    row.realValue = EfuseScriptUtility.ConvertToHex("0b" + onePair.RawData);
                //else
                if (!xBitDefRow.IsDefault)
                {
                    string value = efuseDatalogItem.RawData;
                    xBitDefRow.RealValue = EfuseScriptUtility.ConvertValue(appendRichText, "0b" + (isNeedReversed ? new string([.. value.Reverse()]) : value));
                    string alg = loaderEfuseBitDef.BitDefTable.Rows.FirstOrDefault(x => x[loaderEfuseBitDef.BitDefTable.NameIdx].EqualsIgnoreCase(efuseDatalogItem.Id))![loaderEfuseBitDef.BitDefTable.AlgorithmIdx];
                    if (IdsRegex().IsMatch(alg) &&
                        !MilliampsRegex().IsMatch(efuseDatalogItem.Value))
                    {
                        if (double.TryParse(xBitDefRow.Resolution, out double resolution))
                        {
                            xBitDefRow.RealValue = double.Parse(xBitDefRow.RealValue.ToString()) * resolution;
                        }
                    }

                    if (!EfuseScriptUtility.CheckEfuseLimit(xBitDefRow.RealValue, xBitDefRow.LowLimit, xBitDefRow.HighLimit) && !xBitDefRow.IsDefault)
                    {

                        /*lotid numeric CRC
                     */
                        var exceptionList = new List<string> { "CRC", "lotid", "numeric" };
                        if (exceptionList.All(p => !p.EqualsIgnoreCase(xBitDefRow.Algorithm)))
                        {
                            isFail = true;
                            err += "Out of limit. ";
                        }
                    }
                }
                if (xBitDefRow.IsDefault)
                {
                    #region Check Default Part
                    if (
                        !EfuseScriptUtility.CheckEfuseDefaultValue(appendRichText, efuseDatalogItem.Value, xBitDefRow.DefaultValue))
                    {
                        isFail = true;
                        err += "Should be " +
                               xBitDefRow.DefaultValue + ". ";
                    }
                    #endregion
                }
                else
                {
                    #region Check Real Part
                    double value =
                        EfuseScriptUtility.ConvertValue(appendRichText, efuseDatalogItem.Value);
                    if (value.Equals(0.0) && !ProductIdentifierOrIdsVddRegex().IsMatch(xBitDefRow.Name))
                    {
                        if (xBitDefRow.JobStage <= testCat)
                        {
                            //IsFail = true;
                            err += "Real value should not be 0. ";
                        }
                    }
                    #endregion
                }
                if (isFail)
                {
                    var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_17, EnumErrorLevel.Error, "", row.RowNo, colIdxMain, err);
                    PrintErrorComment(appendRichText, excelWorksheet, error);
                }

                #endregion

            }
        }

        private static void CheckHip(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int colIdxMain, HipItem hipItem, EfuseDatalogItem efuseDatalogItem, XBitDefRow xBitDefRow, string blockname)
        {
            #region Check HIP items

            XBitDefRow bdfItem = XReadEfuseBitDef.GetRow(efuseDatalogItem.Id, blockname)!;
            if (!string.IsNullOrEmpty(bdfItem.HiPparName) &&
                hipItem.HipData.TryGetValue(bdfItem.HiPparName, out string? value))
            {
                excelWorksheet.Cells[xBitDefRow.RowNo, colIdxMain + 1]
                    .Value =
                    value;
                excelWorksheet.Cells[xBitDefRow.RowNo, colIdxMain + 2].Value =
                    bdfItem.HiPparEq;
                EfuseScriptUtility.CheckEquation(appendRichText, excelWorksheet, xBitDefRow.RowNo, colIdxMain);
            }
            #endregion
        }

        private static void PreSetup(XDiceInfo xDiceInfo, ref bool isNeedReversedPrr, DeviceInfo deviceInfo, ref List<string> lotIdList, EfuseDatalogItem efuseDatalogItem, XBitDefRow xBitDefRow)
        {
            if (efuseDatalogItem.Id.EqualsIgnoreCase("x_coordinate"))
            {
                xDiceInfo.PrrX = deviceInfo.XCor = efuseDatalogItem.Value;
                isNeedReversedPrr = int.TryParse(xBitDefRow?.Msb, out int xmsb) && int.TryParse(xBitDefRow?.Lsb, out int xlsb) && xmsb > xlsb;
            }
            else if (efuseDatalogItem.Id.EqualsIgnoreCase("y_coordinate"))
            {
                xDiceInfo.PrrY = deviceInfo.YCor = efuseDatalogItem.Value;
                isNeedReversedPrr = int.TryParse(xBitDefRow?.Msb, out int ymsb) && int.TryParse(xBitDefRow?.Lsb, out int ylsb) && ymsb > ylsb;
            }
            else if (xBitDefRow != null && xBitDefRow.Algorithm.EqualsIgnoreCase("lotid") && xBitDefRow.RealStage.ContainsIgnoreCase("cp1"))
            {
                isNeedReversedPrr = int.TryParse(xBitDefRow?.Msb, out int lmsb) && int.TryParse(xBitDefRow?.Lsb, out int llsb) && lmsb > llsb;
                lotIdList = UpdateLotIdList(isNeedReversedPrr, lotIdList, efuseDatalogItem);
            }
            else if (efuseDatalogItem.Id.EqualsIgnoreCase("wafer_id"))
            {
                xDiceInfo.PrrWaferId = deviceInfo.Wafer = efuseDatalogItem.Value;
            }
        }

        private static List<string> UpdateLotIdList(bool isNeedReversedPrr, List<string> lotIdList, EfuseDatalogItem efuseDatalogItem)
        {
            if (lotIdList.Count != 0)
            {
                if (!isNeedReversedPrr)
                {
                    lotIdList.Insert(0, efuseDatalogItem.Value);
                }
                else
                {
                    lotIdList.Add(efuseDatalogItem.Value);
                }
            }
            else
            {
                lotIdList.Add(efuseDatalogItem.Value);
            }
            // test chip case will contain lot info in all bank
            lotIdList = [.. lotIdList.Distinct()];
            return lotIdList;
        }

        protected override List<XBitDefRow> GetSetWriteValueTable(int testCat)
        {
            return [.. XReadEfuseBitDef.BitDefTable.Where(p => (p.DefaultReal.EqualsIgnoreCase("Real") || p.Algorithm.EqualsIgnoreCase("vddbin")) && !Regex.IsMatch(p.Name) && !p.Algorithm.ContainsIgnoreCase("crc") && p.JobStage == testCat)];
        }

        protected override List<XBitDefRow> GetrealValueTable(int testCat)
        {
            return [.. XReadEfuseBitDef.BitDefTable.Where(p => p.DefaultReal.EqualsIgnoreCase("Real") && !Regex.IsMatch(p.Name) && !p.Algorithm.ContainsIgnoreCase("crc") && p.JobStage == testCat)];
        }
    }
}
