using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using EfuseCheckCmdLib.AlgorithmCheck;
using EfuseCheckCmdLib.IgxlLogLib;
using EfuseCheckCmdLib.Output;
using EfuseCheckCmdLib.Static;
using EfuseCheckCmdLib.Utility;

using LogLib.Utility;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

using TestPlanLib.Efuse;

using EnumErrorLevel = CommonLib.ErrorReport.Base.EnumErrorLevel;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public partial class EFuseAppExcelWriter(EfuseScriptConfig efuseScriptConfig)
    {

        public enum ErrorSheet
        {
            Main,
        }
        private const string FusePattenSummarySheetName = "FusePattenSummary";
        private const string SummaryReportSheetName = "SummaryErrorReport";
        private const string SetWriteVariableSheetName = "SetWriteVariable";
        private const string HipbvRealValueSheetName = "HIPBVRealValue";
        private const string EccCheckSheetName = "EccCheck";
        protected static readonly Regex Regex = IdentifierKeysRegex();

        [GeneratedRegex("PRODUCT_IDENTIFIER|lot_id|wafer_id|x_coordinate|y_coordinate", RegexOptions.IgnoreCase)]
        private static partial Regex IdentifierKeysRegex();

        [GeneratedRegex("UDR", RegexOptions.IgnoreCase)]
        private static partial Regex UdrRegex();

        [GeneratedRegex("CP", RegexOptions.IgnoreCase)]
        private static partial Regex CpRegex();

        [GeneratedRegex("CFG_CONDITION", RegexOptions.IgnoreCase)]
        private static partial Regex CfgConditionExactRegex();

        [GeneratedRegex("PRODUCT_IDENTIFIER|^IDS_VDD", RegexOptions.IgnoreCase)]
        private static partial Regex ProductIdentifierOrIdsVddRegex();

        [GeneratedRegex("ids", RegexOptions.IgnoreCase)]
        private static partial Regex IdsRegex();

        [GeneratedRegex(@"\d\s*ma", RegexOptions.IgnoreCase)]
        private static partial Regex MilliampsRegex();

        [GeneratedRegex("cmp", RegexOptions.IgnoreCase)]
        private static partial Regex CmpRegex();

        [GeneratedRegex("Default", RegexOptions.IgnoreCase)]
        private static partial Regex DefaultRegex();

        [GeneratedRegex("base", RegexOptions.IgnoreCase)]
        private static partial Regex BaseRegex();

        [GeneratedRegex("_WriteDSSC_NV", RegexOptions.IgnoreCase)]
        private static partial Regex WriteDsscNvSuffixRegex();

        [GeneratedRegex("^=", RegexOptions.IgnoreCase)]
        private static partial Regex FormulaPrefixRegex();

        [GeneratedRegex("WLFT", RegexOptions.IgnoreCase)]
        private static partial Regex WlftRegex();

        [GeneratedRegex("^FUSE_REVISION$", RegexOptions.IgnoreCase)]
        private static partial Regex FuseRevisionExactRegex();
        protected EfuseScriptConfig _config = efuseScriptConfig;

        public void WriteMpMainFile(Action<string, string> appendRichText, string filename, LoaderEfuseBitDef loaderEfuseBitDef, List<XParseDatalog.StageDicesInfo> stageDicesInfos, List<XParseDatalog.EFuseSyntaxChkItem> eFuseSyntaxChkItems)
        {
            EfuseCheckResultType result = EfuseCheckResultType.Pass;
            EFuseAppExcelWriterStatic.SummaryErrorsList.Clear();

            FileInfo newFile = Writer.CreateNewFile(filename.Replace(".xlsx", "_MP.xlsx"));
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

        public void WriteReport(Action<string, string> appendRichText, FileInfo fileInfo, LoaderEfuseBitDef loaderEfuseBitDef, List<XParseDatalog.StageDicesInfo> stageDicesInfos, List<XParseDatalog.EFuseSyntaxChkItem> eFuseSyntaxChkItems)
        {
            int testCat = EfuseScriptUtility.GetJobOrder(stageDicesInfos[0].Stage);
            List<XParseDatalog.EFuseSyntaxChkItem> syntaxItems = eFuseSyntaxChkItems;
            var limitLimitation = new Dictionary<string, List<double>>();
            var bdfRealItems = new List<string>();
            var totalCrcItems = new List<List<CrcItem>>();
            var blockListInBdf = loaderEfuseBitDef.BitDefTable.Rows.Select(p => p[loaderEfuseBitDef.BitDefTable.BlockIdx]).Distinct().ToList();
            string runStage = "CP1";
            foreach (string block in blockListInBdf)
            {
                if (block.ContainsIgnoreCase("UDR"))
                {
                    string cmpBlock = UdrRegex().Replace(block, "CMP");
                    _config.UdrcmpMapping[cmpBlock] = block;
                }
            }
            bool isContainsHipInfo = stageDicesInfos.SelectMany(p => p.AlldicesDiceInfos).ToList().Any(p => p.HardIpFuseInfo.Count > 0);
            bool isContainsBvInfo = stageDicesInfos.SelectMany(p => p.AlldicesDiceInfos).ToList().Any(p => p.BvFuseInfo.Count > 0);
            bool isContainsSetWrite = stageDicesInfos.SelectMany(p => p.AlldicesDiceInfos).SelectMany(p => p.HardIpFuseInfo).Any(p => !string.IsNullOrEmpty(p.Value.ReferenceKey));
            bool isContainsEccCheck = LocalSpecs.Options.Device == EnumDevice.RF && stageDicesInfos.SelectMany(p => p.AlldicesDiceInfos).ToList().Any(p => p.EccInfo.Count > 0);
            var realValueTable = new List<XBitDefRow>();
            var setWriteValueTable = new List<XBitDefRow>();
            using var package = new ExcelPackage(fileInfo);
            //STEP1. add a new worksheet to the empty workbook
            ExcelWorksheet summaryWorkSheet = package.Workbook.Worksheets.Add(SummaryReportSheetName);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(SheetConst.Type5BitDefTable);
            ExcelWorksheet hardipWorkSheet = null!;
            ExcelWorksheet setWriteVariableSheet = null!;
            ExcelWorksheet eccCheckSheet = null!;

            //STEP2. Print header
            #region print header
            WriteExcel.PrintHeader(loaderEfuseBitDef, worksheet);
            worksheet.Cells[11, 1].Value = "EfuseCheck";
            //start in 13

            int lastAdRow = 12;
            string curBlock = loaderEfuseBitDef.BitDefTable.Rows.First()[loaderEfuseBitDef.BitDefTable.BlockIdx];
            loaderEfuseBitDef.BitDefTable.Titles[loaderEfuseBitDef.BitDefTable.NameIdx] = curBlock + " eFuse Bit Def";
            for (int colIdx = 0; colIdx < loaderEfuseBitDef.BitDefTable.Titles.Count; colIdx++)
            {
                worksheet.Cells[lastAdRow, colIdx + 1].Value = loaderEfuseBitDef.BitDefTable.Titles[colIdx];
            }

            lastAdRow++;

            XReadEfuseBitDef.BitDefTable.Clear();

            PrintOthers(appendRichText, loaderEfuseBitDef, testCat, syntaxItems, worksheet, ref lastAdRow, ref curBlock);

            #endregion

            if (isContainsHipInfo)
            {
                int rowHeaderindex = 2;
                realValueTable = GetrealValueTable(testCat);
                setWriteValueTable = GetSetWriteValueTable(testCat);
                hardipWorkSheet = PrintSheets(bdfRealItems, isContainsSetWrite, isContainsEccCheck, realValueTable, package, ref setWriteVariableSheet, ref eccCheckSheet, ref rowHeaderindex);
            }

            //STEP3. Print SetWrite & Program data
            PrintSetWrite(appendRichText, loaderEfuseBitDef, stageDicesInfos, ref testCat, limitLimitation, ref bdfRealItems, totalCrcItems, ref runStage, isContainsHipInfo, isContainsBvInfo, realValueTable, setWriteValueTable, package, summaryWorkSheet, ref worksheet, hardipWorkSheet, setWriteVariableSheet, eccCheckSheet);

            CheckIdsLimitWithResolution(appendRichText, package.Workbook, limitLimitation, loaderEfuseBitDef);
            FreezeSheetFormat(package.Workbook, runStage);

            appendRichText("Generating the report .... Please wait for a moment.", "Green");
            package.Save();
        }

        protected virtual void PrintSetWrite(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, List<XParseDatalog.StageDicesInfo> stageDicesInfos, ref int testCat, Dictionary<string, List<double>> limitLimitation, ref List<string> bdfRealItems, List<List<CrcItem>> totalCrcItems, ref string runStage, bool isContainsHipInfo, bool isContainsBvInfo, List<XBitDefRow> realValueTable, List<XBitDefRow> setWriteValueTable, ExcelPackage excelPackage, ExcelWorksheet summaryWorkSheet, ref ExcelWorksheet worksheet, ExcelWorksheet hipBvValueWorkSheet, ExcelWorksheet setWriteVariableSheet, ExcelWorksheet eccCheckSheet)
        {
            int colIdxMain = loaderEfuseBitDef.BitDefTable.Titles.Count + 1;
            int colIdxMainNext = colIdxMain;
            int colIdxConfig = 1;
            int colidxRealValue = 5;
            int colidxSetWrite = 1;
            int eccDataStartCol = 2;
            int count = 0;
            foreach (XParseDatalog.StageDicesInfo allDiceInfo in stageDicesInfos)
            {
                worksheet = excelPackage.Workbook.Worksheets[SheetConst.Type5BitDefTable];
                testCat = EfuseScriptUtility.GetJobOrder(allDiceInfo.Stage);
                runStage = allDiceInfo.Stage;
                int countForDie = 0;
                long idie = 0;
                count++;
                bool isTheLast = count == stageDicesInfos.Count;
                LoopAlldicesDiceInfos(appendRichText, loaderEfuseBitDef, testCat, limitLimitation, ref bdfRealItems, totalCrcItems, runStage, isContainsHipInfo, isContainsBvInfo, realValueTable, setWriteValueTable, excelPackage, worksheet, hipBvValueWorkSheet, setWriteVariableSheet, eccCheckSheet, ref colIdxMain, ref colIdxMainNext, ref colidxRealValue, ref colidxSetWrite, ref eccDataStartCol, allDiceInfo, ref countForDie, ref idie, isTheLast);

                //Create Harvest sheet field sheet
                Writer.SetHarvestSheetSummary(excelPackage, worksheet);
                SetSheetSummaryReport(summaryWorkSheet);
                // If user import Config Table, tool would generate the Config Table Sheet and check with datalog
                // Generally, this flag should be true...
                if (EfuseCfgTableReader1.IsContainCfgTable)
                {
                    //Check CFGTable
                    EfuseCfgTable table = EfuseScriptUtility.CfgTableSel();
                    if (excelPackage.Workbook.Worksheets[table.TableName] == null)
                    {
                        worksheet = excelPackage.Workbook.Worksheets.Add(table.TableName);
                        Writer.PrintCfgTableHeader(appendRichText, worksheet, table, ref colIdxConfig);
                        colIdxConfig++;
                    }
                    worksheet = excelPackage.Workbook.Worksheets[table.TableName];
                    XParseDatalog.TestCat = EfuseScriptUtility.GetJobOrder(allDiceInfo.Stage);

                    foreach (XDiceInfo oneDice in allDiceInfo.AlldicesDiceInfos)
                    {
                        EfuseScriptUtility.CheckCfgCondition(appendRichText, worksheet, oneDice, loaderEfuseBitDef.BitDefTable, table, colIdxConfig);
                        colIdxConfig++;
                    }
                }
            }
        }

        private void LoopAlldicesDiceInfos(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, Dictionary<string, List<double>> limitLimitation, ref List<string> bdfRealItems, List<List<CrcItem>> totalCrcItems, string runStage, bool isContainsHipInfo, bool isContainsBvInfo, List<XBitDefRow> realValueTable, List<XBitDefRow> setWriteValueTable, ExcelPackage excelPackage, ExcelWorksheet worksheet, ExcelWorksheet hipBvValueWorkSheet, ExcelWorksheet setWriteVariableSheet, ExcelWorksheet eccCheckSheet, ref int colIdxMain, ref int colIdxMainNext, ref int colidxRealValue, ref int colidxSetWrite, ref int eccDataStartCol, XParseDatalog.StageDicesInfo stageDicesInfo, ref int countForDie, ref long idie, bool isTheLast)
        {
            foreach (XDiceInfo oneDice in stageDicesInfo.AlldicesDiceInfos)
            {
                string currentBlock = "";
                int offset = 0;
                PreSetup(testCat, oneDice, ref currentBlock, ref offset);
                var deviceRealItems = new List<string>();
                idie++;
                try
                {
                    countForDie++;
                    bool isTheLastForDie = countForDie == stageDicesInfo.AlldicesDiceInfos.Count;
                    appendRichText(oneDice.XCoor + "," + oneDice.YCoor + " => " + (int)(idie * 100 / stageDicesInfo.AlldicesDiceInfos.Count) + "%", "Blue");

                    oneDice.AllReadFromDssc.AddRange(oneDice.AllUdrVer);
                    if (!EfuseScriptUtility.IsPassBin(oneDice.SortBin, _config) && EfuseStatic.ShowType == 1)
                    {
                        continue;
                    }

                    if (EfuseScriptUtility.IsPassBin(oneDice.SortBin, _config) && EfuseStatic.ShowType == 2)
                    {
                        continue;
                    }

                    colIdxMainNext++;

                    PrintHeader(worksheet, colIdxMain, stageDicesInfo, oneDice);

                    CompareXy(appendRichText, oneDice, runStage, worksheet, colIdxMain);

                    var itemList = XReadEfuseBitDef.BitDefTable
                        .Where(p => !p.Algorithm.EqualsIgnoreCase("cond"))
                        .Select(p => p.RowNo).ToList();

                    PrintData(appendRichText, loaderEfuseBitDef, testCat, limitLimitation, ref bdfRealItems, totalCrcItems, isContainsHipInfo, isContainsBvInfo, realValueTable, setWriteValueTable, excelPackage, worksheet, hipBvValueWorkSheet, setWriteVariableSheet, eccCheckSheet, colIdxMain, ref colidxRealValue, ref colidxSetWrite, ref eccDataStartCol, stageDicesInfo, countForDie, isTheLast, oneDice, ref deviceRealItems, isTheLastForDie, itemList);
                }
                catch (Exception e)
                {
                    Console.Write(e.StackTrace);
                }
                foreach (Error generalError in EFuseAppExcelWriterStatic.GeneralCheckerErrors)
                {
                    if (!string.IsNullOrEmpty(generalError.Message))
                    {
                        AddComment(appendRichText, worksheet, generalError);
                    }
                }
                colIdxMain = colIdxMainNext;
            }
        }

        private void PrintData(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, Dictionary<string, List<double>> limitLimitation, ref List<string> bdfRealItems, List<List<CrcItem>> totalCrcItems, bool isContainsHipInfo, bool isContainsBvInfo, List<XBitDefRow> realValueTable, List<XBitDefRow> setWriteValueTable, ExcelPackage excelPackage, ExcelWorksheet worksheet, ExcelWorksheet hipBvValueWorkSheet, ExcelWorksheet setWriteVariableSheet, ExcelWorksheet eccCheckSheet, int colIdxMain, ref int colidxRealValue, ref int colidxSetWrite, ref int eccDataStartCol, XParseDatalog.StageDicesInfo stageDicesInfo, int countForDie, bool isTheLast, XDiceInfo xDiceInfo, ref List<string> deviceRealItems, bool isTheLastForDie, List<int> itemList)
        {
            #region print data
            EFuseAppExcelWriterStatic.GeneralCheckerErrors.Clear();
            bool isNeedIdsCheck = xDiceInfo.IdsFuseInfo.Count > 0;
            //Warning: It's not align calculation with VBT, so tool need reverse for this case
            bool isNeedReversedPrr = true;
            var dsscInfo = new DeviceInfo();
            var lotIdList = new List<string>();

            LoopBitDefTable(appendRichText, loaderEfuseBitDef, testCat, worksheet, colIdxMain, xDiceInfo, itemList, ref isNeedReversedPrr, dsscInfo, ref lotIdList);

            xDiceInfo.PrrLotId = isNeedReversedPrr ? string.Join("", lotIdList.Select(x => x).Reverse()) : string.Join("", lotIdList);
            dsscInfo.Lot = isNeedReversedPrr ? string.Join("", lotIdList.Select(x => x).Reverse()) : string.Join("", lotIdList);
            CheckLotId(colIdxMain, stageDicesInfo, xDiceInfo, dsscInfo);

            if (EfuseCfgTableReader1.IsContainCfgTable)
            {
                EfuseCfgTable table = EfuseScriptUtility.CfgTableSel();
                CheckSvmcFuse(appendRichText, worksheet, xDiceInfo, stageDicesInfo, colIdxMain, table);
            }

            if (isContainsHipInfo || isContainsBvInfo || isNeedIdsCheck)
            {
                deviceRealItems.AddRange(bdfRealItems);
                deviceRealItems = [.. deviceRealItems.Distinct()];
                hipBvValueWorkSheet.Cells[1, colidxRealValue, 1, colidxRealValue + 1].Value = $"X{xDiceInfo.XCoor}_Y{xDiceInfo.YCoor}";
                hipBvValueWorkSheet.Cells[2, colidxRealValue].Value = "Real";
                hipBvValueWorkSheet.Cells[2, colidxRealValue + 1].Value = "eFuse";
                //BDF HIP dsscLog
                const int rowoffset = 3;
                if (isContainsHipInfo)
                {
                    CheckHip(appendRichText, realValueTable, hipBvValueWorkSheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset);
                }

                if (isNeedIdsCheck)
                {
                    CheckIds(appendRichText, realValueTable, hipBvValueWorkSheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset);
                }

                if (isContainsBvInfo)
                {
                    CheckBv(appendRichText, realValueTable, hipBvValueWorkSheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset);
                }
                CheckNotFoundInDatalog(appendRichText, realValueTable, hipBvValueWorkSheet, colidxRealValue, deviceRealItems, rowoffset);

                colidxRealValue += 2;
            }

            if (setWriteVariableSheet != null && xDiceInfo.HardIpFuseInfo.Count > 0 && xDiceInfo.HardIpFuseInfo.Select(p => p.Value.ReferenceValue).Any())
            {
                WriteSetWriteValueSheet(appendRichText, setWriteVariableSheet, setWriteValueTable, xDiceInfo, ref colidxSetWrite, limitLimitation, ref bdfRealItems);
            }

            if (isTheLastForDie && xDiceInfo.FusePatternLines.Count != 0)
            {
                ExcelWorksheet fusePatternSummarySheet = excelPackage.Workbook.Worksheets[FusePattenSummarySheetName] ?? excelPackage.Workbook.Worksheets.Add(FusePattenSummarySheetName);
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
                    CheckAllSiteRealValue(appendRichText, loaderEfuseBitDef, stageDicesInfo.Stage, worksheet, hipBvValueWorkSheet, setWriteVariableSheet);
                }
            }

            HandleNotFoundInDatalog(appendRichText, worksheet, colIdxMain, itemList);

            //Check efuse rule : CRCCheck
            List<CrcItem> crcInfo = CrcCheck(appendRichText, loaderEfuseBitDef, testCat, worksheet, colIdxMain, xDiceInfo);
            totalCrcItems.Add(crcInfo);
            CheckPrr(appendRichText, worksheet, xDiceInfo, colIdxMain, isNeedReversedPrr);
            CheckIedaNull(appendRichText, worksheet, colIdxMain);
            CheckHramEcid(appendRichText, worksheet, xDiceInfo, colIdxMain);
            CheckSvmcFuseSiteFail(appendRichText, worksheet, xDiceInfo, colIdxMain);
            #endregion
        }

        private static void CheckNotFoundInDatalog(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, List<string> deviceRealItems, int rowoffset)
        {
            foreach (string deviceRealItem in deviceRealItems)
            {
                int rowindexReal = xBitDefRows.FindIndex(p => p.Name.EqualsIgnoreCase(deviceRealItem));
                Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);

                var error = new Error(EfuseCheckCmdLibError.E_MissingValue_01, EnumErrorLevel.Error, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, EfuseCheckCmdLibError.E_MissingValue_01.MessageTemplate);
                AddComment(appendRichText, excelWorksheet, error);
            }
        }

        private static void CheckLotId(int colIdxMain, XParseDatalog.StageDicesInfo stageDicesInfo, XDiceInfo xDiceInfo, DeviceInfo deviceInfo)
        {
            if ((xDiceInfo.Prober.XCor != deviceInfo.XCor || xDiceInfo.Prober.YCor != deviceInfo.YCor || xDiceInfo.Prober.Lot != deviceInfo.Lot || xDiceInfo.Prober.Wafer != deviceInfo.Wafer) && CpRegex().IsMatch(stageDicesInfo.Stage))
            {
                var xyError = new Error(EfuseCheckCmdLibError.E_MismatchValue_27, EnumErrorLevel.Error, "", 4, colIdxMain, $"Prober:{xDiceInfo.Prober.Lot}_{xDiceInfo.Prober.Wafer}_{xDiceInfo.Prober.XCor}_{xDiceInfo.Prober.YCor};DSSC:{deviceInfo.Lot}_{deviceInfo.Wafer}_{deviceInfo.XCor}_{deviceInfo.YCor} not same");
                xyError.Comments.Add(xyError.Message);
                EFuseAppExcelWriterStatic.GeneralCheckerErrors.Add(xyError);
            }
        }

        private static void HandleNotFoundInDatalog(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int colIdxMain, List<int> itemList)
        {
            foreach (int rowNo in itemList)
            {
                var missingError = new Error(EfuseCheckCmdLibError.E_MissingValue_01, EnumErrorLevel.Error, "", rowNo, colIdxMain, EfuseCheckCmdLibError.E_MissingValue_01.MessageTemplate);

                PrintErrorComment(appendRichText, excelWorksheet, missingError);
            }
        }

        private void LoopBitDefTable(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, ExcelWorksheet excelWorksheet, int colIdxMain, XDiceInfo xDiceInfo, List<int> itemList, ref bool isNeedReversedPrr, DeviceInfo deviceInfo, ref List<string> lotIdList)
        {
            foreach (XBitDefRow row in XReadEfuseBitDef.BitDefTable)
            {
                if (xDiceInfo.FuseMpDataSet == null)
                {
                    continue;
                }

                EfuseDatalogItem? onePair = xDiceInfo.FuseMpDataSet!.GetData(row.Block, row.Msb, row.Lsb, row.Algorithm, row.Resolution, row.JobStage <= testCat ? EfuseScriptUtility.BaseVoltage.ToString() : "0");
                if (onePair == null)
                {
                    continue;
                }

                onePair.Id = row.Name.Split('#')[1];
                bool isNeedReversed = int.Parse(row.Lsb) > int.Parse(row.Msb);
                PreSetup(xDiceInfo, ref isNeedReversedPrr, deviceInfo, ref lotIdList, row, onePair);

                string blockname = EfuseScriptUtility.CheckAliasBlock(onePair.Block, _config);
                //if (bitDefMain.Contains(idName)) => used to mapping item with Category and Block
                if (row != null)
                {
                    itemList.Remove(row.RowNo);
                    excelWorksheet.Cells[row.RowNo, colIdxMain].Value =
                        row.IsDefault ? EfuseScriptUtility.ConvertValueToSpecifyFormat(appendRichText, onePair.Value, XReadEfuseBitDef.GetType(onePair.Id, blockname))
                            : onePair.Value;

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
                        if (_config.UdrcmpMapping.ContainsKey(blockname))
                        {
                            #region UDR CMP
                            if (!EfuseScriptUtility.CompareUdrcmpItemsMp(appendRichText, blockname, onePair.Id, xDiceInfo, loaderEfuseBitDef, row, testCat, _config))
                            {
                                var checkError1 = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, EnumErrorLevel.Error, "", row.RowNo, colIdxMain, "");
                                EFuseAppExcelWriterStatic.GeneralCheckerErrors.Add(checkError1);
                            }
                            #endregion
                        }
                        else
                        {
                            HandleNonUdrCmp(appendRichText, loaderEfuseBitDef, testCat, excelWorksheet, row, onePair, isNeedReversed, row, colIdxMain);
                        }
                    }
                    else // add to check value is 0 if the items exceed to current stage
                    {
                        CheckIfAfterProgrammingStage(appendRichText, onePair, row, colIdxMain, excelWorksheet);
                    }
                }
            }
        }

        private static void PreSetup(XDiceInfo xDiceInfo, ref bool isNeedReversedPrr, DeviceInfo deviceInfo, ref List<string> lotIdList, XBitDefRow xBitDefRow, EfuseDatalogItem efuseDatalogItem)
        {
            if (efuseDatalogItem.Id.EqualsIgnoreCase("x_coordinate"))
            {
                xDiceInfo.PrrX = efuseDatalogItem.Value;
                deviceInfo.XCor = efuseDatalogItem.Value;
                isNeedReversedPrr = int.Parse(xBitDefRow.Msb) > int.Parse(xBitDefRow.Lsb);
            }
            else if (efuseDatalogItem.Id.EqualsIgnoreCase("y_coordinate"))
            {
                xDiceInfo.PrrY = efuseDatalogItem.Value;
                deviceInfo.YCor = efuseDatalogItem.Value;
                isNeedReversedPrr = int.Parse(xBitDefRow.Msb) > int.Parse(xBitDefRow.Lsb);
            }
            else if (xBitDefRow.Algorithm.EqualsIgnoreCase("lotid") && xBitDefRow.RealStage.ContainsIgnoreCase("cp1"))
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
                isNeedReversedPrr = int.Parse(xBitDefRow.Msb) > int.Parse(xBitDefRow.Lsb);
                lotIdList = [.. lotIdList.Distinct()];
            }
            else if (efuseDatalogItem.Id.EqualsIgnoreCase("wafer_id"))
            {
                xDiceInfo.PrrWaferId = efuseDatalogItem.Value;
                deviceInfo.Wafer = efuseDatalogItem.Value;
                isNeedReversedPrr = int.Parse(xBitDefRow.Msb) > int.Parse(xBitDefRow.Lsb);
            }
        }

        public static void CheckIfAfterProgrammingStage(Action<string, string> appendRichText, EfuseDatalogItem efuseDatalogItem, XBitDefRow xBitDefRow, int colIdxMain, ExcelWorksheet excelWorksheet)
        {
            #region Check item if after programming stage
            string regBin = "0*[bB](?<Number>[01]+)";
            string regHex = "0*[xX](?<Number>[0-9A-F]+)";
            bool isFail = false;
            string value = efuseDatalogItem.Value;
            if (Regex.IsMatch(value, regHex, RegexOptions.IgnoreCase))
            {
                try
                {
                    value = Convert.ToInt64(Regex.Match(value, regHex).Groups["Number"].ToString(), 16).ToString();
                }
                catch (Exception)
                {
                    EfuseStatic.Result = EfuseCheckResultType.Exception;
                }
            }
            if (Regex.IsMatch(value, regBin, RegexOptions.IgnoreCase))
            {
                value =
                    Convert.ToInt64(Regex.Match(value, regBin).Groups["Number"].ToString(), 2).ToString();
            }

            try
            {
                if (double.Parse(value) != 0.0)
                {
                    isFail = true;
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

            if (isFail && !CfgConditionExactRegex().IsMatch(efuseDatalogItem.Id))
            {
                var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_16, EnumErrorLevel.Error, "", xBitDefRow.RowNo, colIdxMain, "Value should be 0. ");
                PrintErrorComment(appendRichText, excelWorksheet, error);
            }
            #endregion
        }

        private static void HandleNonUdrCmp(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, ExcelWorksheet excelWorksheet, XBitDefRow xBitDefRow, EfuseDatalogItem efuseDatalogItem, bool isNeedReversed, XBitDefRow row, int colIdxMain)
        {
            #region non UDR CMP items
            bool isFail = false;
            string err = "Fail: ";
            if (!xBitDefRow.IsDefault)
            {
                HandleBitDefRow(appendRichText, loaderEfuseBitDef, xBitDefRow, efuseDatalogItem, isNeedReversed, ref isFail, ref err);
            }
            if (xBitDefRow.IsDefault)
            {
                #region Check Default Part
                if (
                    !EfuseScriptUtility.CheckEfuseDefaultValue(appendRichText, efuseDatalogItem.Value, xBitDefRow.DefaultValue))
                {
                    isFail = true;
                    err += "Should be " + xBitDefRow.DefaultValue + ". ";
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

                var checkError = new Error(EfuseCheckCmdLibError.E_MismatchValue_17, EnumErrorLevel.Error, "", row.RowNo, colIdxMain, err);
                PrintErrorComment(appendRichText, excelWorksheet, checkError);
            }
            #endregion
        }

        private static void HandleBitDefRow(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, XBitDefRow xBitDefRow, EfuseDatalogItem efuseDatalogItem, bool isNeedReversed, ref bool isFail, ref string err)
        {
            xBitDefRow.RealValue = EfuseScriptUtility.ConvertValue(appendRichText, "0b" + (isNeedReversed ? efuseDatalogItem.RawData.Reverse() : efuseDatalogItem.RawData));
            string alg = loaderEfuseBitDef.BitDefTable.Rows.FirstOrDefault(x => x[loaderEfuseBitDef.BitDefTable.NameIdx].EqualsIgnoreCase(efuseDatalogItem.Id))![loaderEfuseBitDef.BitDefTable.AlgorithmIdx];
            if (IdsRegex().IsMatch(alg) &&
                !MilliampsRegex().IsMatch(efuseDatalogItem.Value))
            {
                if (double.TryParse(xBitDefRow.Resolution, out double resolution))
                {
                    xBitDefRow.RealValue = double.Parse(xBitDefRow.RealValue.ToString()) * resolution;
                }
            }

            CheckEfuseLimit(xBitDefRow, ref isFail, ref err);
        }

        private static void CheckEfuseLimit(XBitDefRow xBitDefRow, ref bool isFail, ref string err)
        {
            if (!EfuseScriptUtility.CheckEfuseLimit(xBitDefRow.RealValue, xBitDefRow.LowLimit, xBitDefRow.HighLimit) && !xBitDefRow.IsDefault)
            {
                // lotid numeric CRC
                var exceptionList = new List<string> { "CRC", "lotid", "numeric" };
                if (exceptionList.All(p => !p.EqualsIgnoreCase(xBitDefRow.Algorithm)))
                {
                    isFail = true;
                    err += "Out of limit. ";
                }
            }
        }

        public static List<CrcItem> CrcCheck(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, ExcelWorksheet excelWorksheet, int colIdxMain, XDiceInfo xDiceInfo)
        {
            var crcInfo = new List<CrcItem>();
            foreach (CrcItem crc in loaderEfuseBitDef.BitDefTable.CrcItem)
            {
                if (EfuseScriptUtility.GetJobOrder(crc.Job) <= testCat)
                {
                    EfuseScriptUtility.CheckCrc(appendRichText, excelWorksheet, crc, xDiceInfo, loaderEfuseBitDef.BitDefTable, colIdxMain);
                    crcInfo.Add(crc);
                }
            }

            return crcInfo;
        }

        public static void CheckBv(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, List<string> deviceRealItems, int rowoffset)
        {
            #region BV Check
            var bvDsscItems = xDiceInfo.AllReadFromDssc.Where(p => xDiceInfo.BvFuseInfo.Keys.Any(
q => q.EqualsIgnoreCase(p.Id.ToLower()) && !CmpRegex().IsMatch(p.Block))).ToList();

            foreach (EfuseDatalogItem bvDsscItem in bvDsscItems)
            {
                try
                {
                    deviceRealItems.Remove(bvDsscItem.Block + "#" + bvDsscItem.Id.ToUpper());
                    int rowindexReal =
                        xBitDefRows.FindIndex(
                            p =>
                                p.Name.EqualsIgnoreCase(bvDsscItem.Block + "#" + bvDsscItem.Id));
                    string hipTarget =
                        xDiceInfo.BvFuseInfo.Keys.FirstOrDefault(
                            p =>
                                p.EqualsIgnoreCase(bvDsscItem.Id))!;
                    if (rowindexReal != -1)
                    {
                        excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue].Value
                            =
                            xDiceInfo.BvFuseInfo[hipTarget];
                        excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue + 1]
                            .Value
                            = bvDsscItem.Value;
                        if (double.Parse(xDiceInfo.BvFuseInfo[hipTarget]) != double.Parse(bvDsscItem.Value))
                        {
                            var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, EnumErrorLevel.Error, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, "Compared item not match.");
                            AddComment(appendRichText, excelWorksheet, error);

                            Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);
                        }

                    }
                    else
                    {
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
            #endregion
        }

        public static void CheckIds(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, List<string> deviceRealItems, int rowoffset)
        {
            #region IDS Check

            var idsDsscItems = xDiceInfo.AllReadFromDssc.Where
                (p =>
                    xDiceInfo.IdsFuseInfo.Keys.Any(
                        q => p.Id.EqualsIgnoreCase(q.ToLower().Split('#')[1]) &&
                             p.Block.EqualsIgnoreCase(q.ToLower().Split('#')[0]))).ToList();

            foreach (EfuseDatalogItem idsDsscItem in idsDsscItems)
            {
                deviceRealItems.Remove(idsDsscItem.Block + "#" + idsDsscItem.Id.ToUpper());
                try
                {
                    int rowindexReal =
                        xBitDefRows.FindIndex(
                            p =>
                                p.Name.EqualsIgnoreCase(idsDsscItem.Block + "#" + idsDsscItem.Id));
                    string idsMeasKey =
                        xDiceInfo.IdsFuseInfo.Keys.FirstOrDefault(
                            p =>
                                p.EqualsIgnoreCase(idsDsscItem.Block + "#" + idsDsscItem.Id))!;
                    if (rowindexReal != -1)
                    {
                        excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue].Value
                            =
                            xDiceInfo.IdsFuseInfo[idsMeasKey];
                        excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue + 1]
                            .Value
                            = idsDsscItem.Value;
                        double resolution = double.Parse(xBitDefRows[rowindexReal].Resolution);
                        double idsMeasValue = double.Parse(xDiceInfo.IdsFuseInfo[idsMeasKey]);
                        double expectIdsRealValue = EfuseStatic.IdsRoundMethod ? Math.Floor(idsMeasValue / resolution) * resolution :
                            Math.Ceiling(idsMeasValue / resolution) * resolution;
                        if (Math.Abs(expectIdsRealValue - double.Parse(idsDsscItem.Value)) > resolution)
                        {
                            var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_25, EnumErrorLevel.Error, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, EfuseCheckCmdLibError.E_MismatchValue_25.MessageTemplate);
                            AddComment(appendRichText, excelWorksheet, error);

                            Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);
                        }

                    }
                    else
                    {
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
            #endregion
        }

        private static void CheckHip(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, List<string> deviceRealItems, int rowoffset)
        {
            #region HIP Check
            var hipDsscItems = xDiceInfo.AllReadFromDssc.Where(p => xDiceInfo.HardIpFuseInfo.Keys.Any(q => p.Id.EqualsIgnoreCase(q.ToLower().Split('#')[1]) && p.Block.EqualsIgnoreCase(q.ToLower().Split('#')[0]))).ToList();
            foreach (EfuseDatalogItem hipDsscItem in hipDsscItems)
            {
                CheckOneHipRow(appendRichText, xBitDefRows, excelWorksheet, colidxRealValue, xDiceInfo, deviceRealItems, rowoffset, hipDsscItem);
            }
            #endregion
        }

        private static void CheckOneHipRow(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, List<string> deviceRealItems, int rowoffset, EfuseDatalogItem efuseDatalogItem)
        {
            deviceRealItems.Remove(efuseDatalogItem.Block + "#" + efuseDatalogItem.Id.ToUpper());
            try
            {
                int rowindexReal = xBitDefRows.FindIndex(p => p.Name.EqualsIgnoreCase(efuseDatalogItem.Block + "#" + efuseDatalogItem.Id));
                string hipTarget = xDiceInfo.HardIpFuseInfo.Keys.FirstOrDefault(p => p.EqualsIgnoreCase(efuseDatalogItem.Block + "#" + efuseDatalogItem.Id))!;
                if (rowindexReal != -1)
                {
                    excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue].Value = xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.TrimStart().Split(' ')[0];
                    excelWorksheet.Cells[rowindexReal + rowoffset, colidxRealValue + 1].Value = efuseDatalogItem.Value;
                    if (xBitDefRows[rowindexReal].Algorithm.Trim().EqualsIgnoreCase("ids"))
                    {
                        CheckAlgorithmIsIds(appendRichText, xBitDefRows, excelWorksheet, colidxRealValue, xDiceInfo, rowoffset, efuseDatalogItem, rowindexReal, hipTarget);
                    }
                    else
                    {
                        if (!xBitDefRows[rowindexReal].Algorithm.Trim().EqualsIgnoreCase("vddbin"))
                        {
                            CheckAlgorithmNotVddbin(appendRichText, excelWorksheet, colidxRealValue, xDiceInfo, rowoffset, efuseDatalogItem, rowindexReal, hipTarget);
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

        private static void CheckAlgorithmNotVddbin(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, int rowoffset, EfuseDatalogItem efuseDatalogItem, int rowindexReal, string hipTarget)
        {
            string value = xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.Split(' ')[0];
            value = value.IsHexValue() ? Convert.ToInt64(value, 16).ToString() : value;
            string hipDsscValue = efuseDatalogItem.Value.IsHexValue() ? Convert.ToInt64(efuseDatalogItem.Value, 6).ToString() : efuseDatalogItem.Value;
            if (xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.Split(' ')[0] != hipDsscValue)
            {
                var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, EnumErrorLevel.Error, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, "Compared item not match.");
                AddComment(appendRichText, excelWorksheet, error);
                Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);
            }
        }

        private static void CheckAlgorithmIsIds(Action<string, string> appendRichText, List<XBitDefRow> xBitDefRows, ExcelWorksheet excelWorksheet, int colidxRealValue, XDiceInfo xDiceInfo, int rowoffset, EfuseDatalogItem efuseDatalogItem, int rowindexReal, string hipTarget)
        {
            string regIds = @"\(\s*(?<value>[\-\w\.]+)";
            string idsValue = Regex.Match(xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue, regIds, RegexOptions.IgnoreCase).Groups["value"].Value;
            double resolution = double.Parse(xBitDefRows[rowindexReal].Resolution);
            double idsMeasValue = double.Parse(idsValue);
            double expectIdsRealValue = EfuseStatic.IdsRoundMethod ? Math.Floor(idsMeasValue / resolution) * resolution : Math.Ceiling(idsMeasValue / resolution) * resolution;
            double tempValue = double.Parse(idsValue) / resolution;
            if (Math.Abs(expectIdsRealValue - double.Parse(efuseDatalogItem.Value)) > resolution && Math.Abs(tempValue - double.Parse(xDiceInfo.HardIpFuseInfo[hipTarget].ReferenceValue.TrimStart().Split(' ')[0])) > resolution)
            {
                var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_26, EnumErrorLevel.Error, excelWorksheet.Name, rowindexReal + rowoffset, colidxRealValue, EfuseCheckCmdLibError.E_MismatchValue_26.MessageTemplate);
                AddComment(appendRichText, excelWorksheet, error);
                Writer.HighLightCell(excelWorksheet, rowindexReal + rowoffset, colidxRealValue, rowindexReal + rowoffset, colidxRealValue + 1, Color.Red);
            }
        }

        public static void PrintHeader(ExcelWorksheet excelWorksheet, int colIdxMain, XParseDatalog.StageDicesInfo stageDicesInfo, XDiceInfo xDiceInfo)
        {
            excelWorksheet.Cells[1, colIdxMain].Value = xDiceInfo.Sort;
            excelWorksheet.Cells[2, colIdxMain].Value = xDiceInfo.SortBin;
            excelWorksheet.Cells[3, colIdxMain].Value = stageDicesInfo.Stage;
            //Print XY Coordinate 
            excelWorksheet.Cells[4, colIdxMain].Value = $"X{xDiceInfo.XCoor}_Y{xDiceInfo.YCoor}";

            excelWorksheet.Cells[5, colIdxMain].Value = xDiceInfo.PrrCode;

            //Print efuse info
            excelWorksheet.Cells[6, colIdxMain].Value = xDiceInfo.EFuseLotNumber;
            excelWorksheet.Cells[7, colIdxMain].Value = xDiceInfo.EFuseWaferId;
            excelWorksheet.Cells[8, colIdxMain].Value = xDiceInfo.EFuseDieX;
            excelWorksheet.Cells[9, colIdxMain].Value = xDiceInfo.EFuseDieY;
            excelWorksheet.Cells[10, colIdxMain].Value = xDiceInfo.HramEcid53Bit;
            excelWorksheet.Cells[11, colIdxMain].Value = xDiceInfo.SvmCFuse288Bits;
        }

        private static void PreSetup(int testCat, XDiceInfo xDiceInfo, ref string currentBlock, ref int offset)
        {
            if (xDiceInfo.AllReadFromDssc.Count == 0 && xDiceInfo.FuseMpDataSet != null)
            {
                foreach (XBitDefRow row in XReadEfuseBitDef.BitDefTable)
                {
                    if (string.IsNullOrEmpty(currentBlock) || currentBlock != row.Block)
                    {
                        currentBlock = row.Block;
                        offset = int.Parse(row.Msb) > int.Parse(row.Lsb) ? int.Parse(row.Lsb) : int.Parse(row.Msb);
                    }
                    int intMsb = int.Parse(row.Msb) - offset;
                    int intLsb = int.Parse(row.Lsb) - offset;
                    row.Msb = intMsb.ToString();
                    row.Lsb = intLsb.ToString();
                    EfuseDatalogItem? onePair = xDiceInfo.FuseMpDataSet!.GetData(row.Block, row.Msb, row.Lsb, row.Algorithm, row.Resolution, row.JobStage <= testCat ? EfuseScriptUtility.BaseVoltage.ToString() : "0");
                    if (onePair == null)
                    {
                        continue;
                    }

                    onePair.Id = row.Name.Split('#')[1];
                    xDiceInfo.AllReadFromDssc.Add(onePair);
                }
            }
        }

        protected virtual List<XBitDefRow> GetSetWriteValueTable(int testCat)
        {
            return [.. XReadEfuseBitDef.BitDefTable.Where(p => p.DefaultReal.EqualsIgnoreCase("Real") && !Regex.IsMatch(p.Name) && p.JobStage == testCat)];
        }

        protected virtual List<XBitDefRow> GetrealValueTable(int testCat)
        {
            return [.. XReadEfuseBitDef.BitDefTable.Where(p => (p.DefaultReal.EqualsIgnoreCase("Real") || p.Algorithm.EqualsIgnoreCase("vddbin") || p.Algorithm.EqualsIgnoreCase("ids")) && !Regex.IsMatch(p.Name) && !p.Algorithm.EqualsIgnoreCase("crc") && p.JobStage == testCat)];
        }

        public static ExcelWorksheet PrintSheets(List<string> bdfRealItems, bool isContainsSetWrite, bool isContainsEccCheck, List<XBitDefRow> xBitDefRows, ExcelPackage excelPackage, ref ExcelWorksheet setWriteVariableSheet, ref ExcelWorksheet eccCheckSheet, ref int rowHeaderindex)
        {
            ExcelWorksheet hardipWorkSheet;
            if (isContainsSetWrite)
            {
                setWriteVariableSheet = excelPackage.Workbook.Worksheets.Add(SetWriteVariableSheetName);
            }
            if (isContainsEccCheck)
            {
                eccCheckSheet = excelPackage.Workbook.Worksheets.Add(EccCheckSheetName);
            }
            hardipWorkSheet = excelPackage.Workbook.Worksheets.Add(HipbvRealValueSheetName);
            hardipWorkSheet.Cells[rowHeaderindex, 1].Value = "Category";
            hardipWorkSheet.Cells[rowHeaderindex, 2].Value = "Bank";
            hardipWorkSheet.Cells[rowHeaderindex, 3].Value = "Stage";
            hardipWorkSheet.Cells[rowHeaderindex, 4].Value = "DefaultReal";
            rowHeaderindex++;
            foreach (XBitDefRow categoryInfo in xBitDefRows)
            {
                bdfRealItems.Add(categoryInfo.Name);
                hardipWorkSheet.Cells[rowHeaderindex, 1].Value = categoryInfo.Name.Split('#')[1];
                hardipWorkSheet.Cells[rowHeaderindex, 2].Value = categoryInfo.Name.Split('#')[0];
                hardipWorkSheet.Cells[rowHeaderindex, 3].Value = categoryInfo.RealStage;
                hardipWorkSheet.Cells[rowHeaderindex, 4].Value = categoryInfo.DefaultReal;
                rowHeaderindex++;
            }

            return hardipWorkSheet;
        }

        protected virtual void PrintOthers(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, int testCat, List<XParseDatalog.EFuseSyntaxChkItem> eFuseSyntaxChkItems, ExcelWorksheet excelWorksheet, ref int lastAdRow, ref string curBlock)
        {
            foreach (List<string> oneRow in loaderEfuseBitDef.BitDefTable.Rows)
            {
                if (curBlock != oneRow[loaderEfuseBitDef.BitDefTable.BlockIdx])
                {
                    // skip 1 row if items of block is different
                    lastAdRow++;
                    Writer.PrintBlockHeader(excelWorksheet, loaderEfuseBitDef.BitDefTable, oneRow, ref curBlock, ref lastAdRow);
                    //curBlock = oneRow[this.bitDefRef.bitDefTable.blockIdx];
                }
                var oneBitRow = new XBitDefRow { Block = oneRow[loaderEfuseBitDef.BitDefTable.BlockIdx] };
                oneBitRow.Name = oneBitRow.Block + "#" + oneRow[loaderEfuseBitDef.BitDefTable.NameIdx];
                oneBitRow.Resolution = oneRow[loaderEfuseBitDef.BitDefTable.ResolutionIdx];
                oneBitRow.RowNo = lastAdRow;
                oneBitRow.DefaultReal = oneRow[loaderEfuseBitDef.BitDefTable.DefaultOrRealIdx];
                oneBitRow.IsDefault = DefaultRegex().IsMatch(oneRow[loaderEfuseBitDef.BitDefTable.DefaultOrRealIdx]);

                oneBitRow.DefaultValue = oneBitRow.IsDefault
                    ? oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx]
                    : "-999";

                oneBitRow.Type = oneBitRow.IsDefault
                    ? EfuseScriptUtility.JudgeValueType(oneRow[loaderEfuseBitDef.BitDefTable.DefaultValueIdx])
                    : EnumValueType.Other;
                if (!_config.UdrcmpMapping.ContainsKey(oneBitRow.Block) && oneBitRow.Block.Contains("CMP"))
                {
                    _config.UdrcmpMapping.Add(oneBitRow.Block, oneBitRow.Block.Replace("CMP", "UDR"));
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

                //if (!oneBitRow.block.Contains("CMP"))  //udr/uid/config...
                //{
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
                        XReadEfuseBitDef.BitDefTable.Add(oneBitRow);
                    }
                }

                lastAdRow++;
            }
        }

        public static void CompareXy(Action<string, string> appendRichText, XDiceInfo xDiceInfo, string currentStage, ExcelWorksheet excelWorksheet, int colIdxMain)
        {
            //for WLFT should get the info_x_coo, CP & FT should get x_coo
            bool isInfoXy = currentStage.StartsWith("WLFT");
            EfuseDatalogItem? xoo = xDiceInfo.AllReadFromDssc.FirstOrDefault(x => isInfoXy ? x.Id.ContainsIgnoreCase("_x_coo") : x.Id.StartsWith("x_coo"));
            EfuseDatalogItem? yoo = xDiceInfo.AllReadFromDssc.FirstOrDefault(x => isInfoXy ? x.Id.ContainsIgnoreCase("_y_coo") : x.Id.StartsWith("y_coo"));
            if (xoo == null || yoo == null)
            {
                var error = new Error(EfuseCheckCmdLibError.E_MissingValue_04, EnumErrorLevel.Error, excelWorksheet.Name, 4, colIdxMain, "Can not find x_coo or y_coo");
                PrintError(appendRichText, excelWorksheet, error, Color.Red);
            }
            else
            {
                if (!int.TryParse(xoo.Value, out int parsedX) || !int.TryParse(yoo.Value, out int parsedY))
                {
                    var error = new Error(EfuseCheckCmdLibError.E_MissingValue_05, EnumErrorLevel.Error, excelWorksheet.Name, 4, colIdxMain, $"X or Y value not found from dssc value, x: {xoo.Value}, y: {yoo.Value}");
                    PrintError(appendRichText, excelWorksheet, error, Color.Red);
                }
                else if (parsedX != xDiceInfo.XCoor || parsedY != xDiceInfo.YCoor)
                {
                    var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_32, EnumErrorLevel.Error, excelWorksheet.Name, 4, colIdxMain, $"X{xDiceInfo.XCoor}_Y{xDiceInfo.YCoor}, from dssc value (x_coo, y_coo):X{parsedX}_Y{parsedY}");
                    PrintError(appendRichText, excelWorksheet, error, Color.Red);
                }
            }
        }

        public static void CheckAllSiteRealValue(Action<string, string> appendRichText, LoaderEfuseBitDef loaderEfuseBitDef, string currectStage, ExcelWorksheet? bdfWorksheet, ExcelWorksheet? hipWorksheet, ExcelWorksheet? setWriteVariableSheet)
        {
            if (bdfWorksheet != null)
            {
                EfuseBitDefTable bdfTable = loaderEfuseBitDef.BitDefTable;
                for (int startRow = 7; startRow <= bdfWorksheet.Dimension.Rows; startRow++)
                {
                    bool isFirst = true;
                    bool isAllSame = false;
                    string firstVal = "";
                    //conditions
                    //1. Currect Job
                    //2. Real value
                    //3. Alg is not cond and CRC, IDS
                    object defaultOrReal = bdfWorksheet.Cells[startRow, bdfTable.DefaultOrRealIdx + 1].Value;
                    object algValue = bdfWorksheet.Cells[startRow, bdfTable.AlgorithmIdx + 1].Value;
                    object programmingStage = bdfWorksheet.Cells[startRow, bdfTable.PrgStageIdx + 1].Value;

                    if (programmingStage != null && programmingStage.ToString()!.EqualsIgnoreCase(currectStage) && defaultOrReal != null && defaultOrReal.ToString()!.EqualsIgnoreCase("real") && algValue != null && !algValue.ToString()!.EqualsIgnoreCase("cond") && !algValue.ToString()!.EqualsIgnoreCase("crc") && !algValue.ToString()!.EqualsIgnoreCase("ids"))
                    {

                        for (int startCol = bdfTable.AccessModeIdx + 2; startCol <= bdfWorksheet.Dimension.Columns; startCol++)
                        {
                            if (bdfWorksheet.Cells[startRow, startCol].Value == null)
                            {
                                continue;
                            }

                            string valStr = bdfWorksheet.Cells[startRow, startCol].Value.ToString()!;
                            if (isFirst)
                            {
                                firstVal = valStr.IsHexValue() ? Convert.ToInt64(valStr, 16).ToString() : valStr;
                                isFirst = false;
                                continue;
                            }
                            else
                            {
                                isAllSame = firstVal == (valStr.IsHexValue() ? Convert.ToInt64(valStr, 16).ToString() : valStr);
                                if (!isAllSame)
                                {
                                    break;
                                }
                            }

                        }
                        if (isAllSame)
                        {
                            var error = new Error(EfuseCheckCmdLibError.I_MismatchValue_01, EnumErrorLevel.Info, bdfWorksheet.Name, startRow, 1, EfuseCheckCmdLibError.I_MismatchValue_01.MessageTemplate);
                            AddComment(appendRichText, bdfWorksheet, error);
                        }
                    }
                }
            }

            if (hipWorksheet != null)
            {
                CheckSetWriteVariableSheet(appendRichText, hipWorksheet, 5);
            }

            if (setWriteVariableSheet != null)
            {
                CheckSetWriteVariableSheet(appendRichText, setWriteVariableSheet, 4);
            }
        }

        private static void CheckSetWriteVariableSheet(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int col)
        {
            for (int startRow = 3; startRow <= excelWorksheet.Dimension.Rows; startRow++)
            {
                bool isFirst = true;
                bool isAllSame = false;
                string firstVal = "";
                for (int startCol = col; startCol <= excelWorksheet.Dimension.Columns; startCol += 2)
                {
                    if (excelWorksheet.Cells[startRow, startCol].Value == null)
                    {
                        continue;
                    }

                    string valStr = excelWorksheet.Cells[startRow, startCol].Value.ToString()!;
                    if (isFirst)
                    {
                        firstVal = valStr.IsHexValue()
                                                            ? Convert.ToInt64(valStr, 16).ToString()
                                                            : valStr;
                        isFirst = false;
                        continue;
                    }
                    else
                    {
                        isAllSame = firstVal == (valStr.IsHexValue()
                                                            ? Convert.ToInt64(valStr, 16).ToString()
                                                            : valStr);
                        if (!isAllSame)
                        {
                            break;
                        }
                    }

                }
                if (isAllSame)
                {
                    var error = new Error(EfuseCheckCmdLibError.I_MismatchValue_01, EnumErrorLevel.Info, excelWorksheet.Name, startRow, 1, EfuseCheckCmdLibError.I_MismatchValue_01.MessageTemplate);
                    AddComment(appendRichText, excelWorksheet, error);
                }
            }
        }

        public static void SetSheetSummaryReport(ExcelWorksheet excelWorksheet)
        {
            #region Print header
            excelWorksheet.Cells[1, 1].Value = "Summary";
            excelWorksheet.Cells[2, 1].Value = "Error Sheet";
            excelWorksheet.Cells[2, 2].Value = "Error Count";
            Writer.HighLightCell(excelWorksheet, 2, 1, 2, 6, Color.Gray);

            #region Print summary count
            var groupDict = EFuseAppExcelWriterStatic.SummaryErrorsList.GroupBy(p => p.SheetName).ToDictionary(grp => grp.Key, grp => grp);

            #endregion

            excelWorksheet.Cells[9, 1].Value = "SheetName";
            excelWorksheet.Cells[9, 2].Value = "Level";
            excelWorksheet.Cells[9, 3].Value = "Link";
            excelWorksheet.Cells[9, 4].Value = "Row";
            excelWorksheet.Cells[9, 5].Value = "Col";
            excelWorksheet.Cells[9, 6].Value = "ErrorMessage";
            Writer.HighLightCell(excelWorksheet, 9, 1, 9, 6, Color.Gray);
            #endregion

            int startErrRow = 3;

            foreach (KeyValuePair<string, IGrouping<string, Error>> error in groupDict)
            {
                excelWorksheet.Cells[startErrRow, 1].Value = error.Key;
                excelWorksheet.Cells[startErrRow, 2].Value = error.Value.Count();
                startErrRow += 1;
            }

            startErrRow = 10;

            foreach (string sheetName in groupDict.Keys)
            {
                IGrouping<string, Error> errorList = groupDict[sheetName];
                foreach (Error error in errorList)
                {
                    excelWorksheet.Cells[startErrRow, 1].Value = error.SheetName;
                    excelWorksheet.Cells[startErrRow, 2].Value = error.ErrorLevel == EnumErrorLevel.Error ? EnumErrorLevel.Error : EnumErrorLevel.Warning;
                    excelWorksheet.Cells[startErrRow, 3].Hyperlink = new ExcelHyperLink(error.SheetName + "!" + error.GetAddress(), "Link");
                    excelWorksheet.Cells[startErrRow, 3].StyleName = "HyperLink";
                    excelWorksheet.Cells[startErrRow, 3].Style.Font.Color.SetColor(Color.Blue);
                    excelWorksheet.Cells[startErrRow, 3].Style.Font.UnderLine = true;
                    excelWorksheet.Cells[startErrRow, 4].Value = error.RowNum;
                    excelWorksheet.Cells[startErrRow, 5].Value = error.ColLetter;
                    excelWorksheet.Cells[startErrRow, 6].Value = error.Message;
                    Writer.HighLightCell(excelWorksheet, startErrRow, 2, startErrRow, 2, Color.Red);
                    startErrRow += 1;
                }
            }
        }

        public static void CheckIdsLimitWithResolution(Action<string, string> appendRichText, ExcelWorkbook excelWorkbook, Dictionary<string, List<double>> limitLimitation, LoaderEfuseBitDef loaderEfuseBitDef)
        {
            ExcelWorksheet worksheet = excelWorkbook.Worksheets[SheetConst.Type5BitDefTable];
            foreach (KeyValuePair<string, List<double>> idsLimit in limitLimitation)
            {
                XBitDefRow row = XReadEfuseBitDef.GetRow(idsLimit.Key.Split('#')[1], idsLimit.Key.Split('#')[0])!;
                double idsSum = idsLimit.Value.Sum();
                if (idsLimit.Value.Count > 1 && idsSum > row.HighLimit * 0.001) //Multiple pins show warning message
                {
                    var missingError = new Error(EfuseCheckCmdLibError.W_MismatchValue_03, EnumErrorLevel.Warning, "", row.RowNo, loaderEfuseBitDef.BitDefTable.HighLimitIdx + 1, $"{idsLimit.Key.Split('#')[1]}:Sum of related pin limit value exceed table range");

                    PrintWarningComment(appendRichText, worksheet, missingError);
                }
                else if (idsSum > row.HighLimit * 0.001)
                {
                    var missingError = new Error(EfuseCheckCmdLibError.W_MismatchValue_04, EnumErrorLevel.Warning, "", row.RowNo, loaderEfuseBitDef.BitDefTable.HighLimitIdx + 1, $"{idsLimit.Key.Split('#')[1]}:Meas limit:{idsLimit.Value.First()} exceed table");

                    PrintWarningComment(appendRichText, worksheet, missingError);
                }
            }
        }

        public static void WriteSetWriteValueSheet(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, List<XBitDefRow> xBitDefRows, XDiceInfo xDiceInfo, ref int colIndex, Dictionary<string, List<double>> idsLimits, ref List<string> bdfRealItems)
        {
            try
            {
                int rowindex = 2;
                List<SetHipDicItem> measLines = xDiceInfo.HipMeasLines;
                List<DataFormatDataRow> idsLines = xDiceInfo.IdsMeasLines;
                excelWorksheet.Cells[2, 1].Value = "Dictionary Name";
                excelWorksheet.Cells[2, 2].Value = "Block";
                excelWorksheet.Cells[2, 3].Value = "Reference";
                rowindex++;

                if (colIndex < 4)
                {
                    colIndex = 4;
                }

                rowindex = WriteData(appendRichText, excelWorksheet, xBitDefRows, xDiceInfo, colIndex, idsLimits, bdfRealItems, rowindex, measLines, idsLines);
                colIndex += 2;
            }
            catch (Exception)
            {
                EfuseStatic.Result = EfuseCheckResultType.Exception;
            }
        }

        private static int WriteData(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, List<XBitDefRow> xBitDefRows, XDiceInfo xDiceInfo, int colIndex, Dictionary<string, List<double>> idsLimits, List<string> bdfRealItems, int rowindex, List<SetHipDicItem> setHipDicItems, List<DataFormatDataRow> dataFormatDataRows)
        {
            foreach (KeyValuePair<string, SetWriteItem> setWriteItem in xDiceInfo.HardIpFuseInfo)
            {
                XBitDefRow? bdfItem = xBitDefRows.FirstOrDefault(p => p.Name.EqualsIgnoreCase(setWriteItem.Key));
                if (bdfItem != null && !string.IsNullOrEmpty(setWriteItem.Value.ReferenceKey))
                {
                    bdfRealItems.Add(bdfItem.Name);

                    excelWorksheet.Cells[1, colIndex].Value = $"X{xDiceInfo.XCoor}_Y{xDiceInfo.YCoor}";
                    excelWorksheet.Cells[2, colIndex].Value = "Fuse Real Value";
                    excelWorksheet.Cells[2, colIndex + 1].Value = "Meas Value";

                    excelWorksheet.Cells[rowindex, 1].Value = bdfItem.Name.Split('#')[1];
                    excelWorksheet.Cells[rowindex, 2].Value = bdfItem.Name.Split('#')[0];
                    excelWorksheet.Cells[rowindex, 3].Value = setWriteItem.Value.ReferenceKey;

                    EfuseDatalogItem? value2 =
                        xDiceInfo.AllReadFromDssc.FirstOrDefault(
                            p => p.Block.EqualsIgnoreCase(setWriteItem.Key.Split('#')[0]) &&
                                 p.Id.EqualsIgnoreCase(setWriteItem.Key.Split('#')[1]) &&
                                 xDiceInfo.Site.ToString() == setWriteItem.Value.Site);
                    if (value2 != null)
                    {
                        try
                        {
                            excelWorksheet.Cells[rowindex, colIndex].Value = Convert.ToInt32(value2.RawData, 2);
                        }
                        catch (Exception)
                        {
                            EfuseStatic.Result = EfuseCheckResultType.Exception;
                        }
                    }
                    //IDS
                    string alg = xBitDefRows.FirstOrDefault(x => x.Name.EqualsIgnoreCase(setWriteItem.Key))!.Algorithm;
                    if (IdsRegex().IsMatch(alg))
                    {

                        var valueList = new List<double>();
                        var limitList = new List<double>();
                        bool keyFound = true;
                        foreach (string refKey in setWriteItem.Value.ReferenceKey.Split(','))
                        {
                            var idsItems = dataFormatDataRows.Where(p => p.Pin.EqualsIgnoreCase(refKey) &&
                                                                p.ActiveSite.ToString() == setWriteItem.Value.Site
                                                                && p.ActuallineNumber < setWriteItem.Value.Number)
                                .ToList();
                            if (idsItems.Count == 0)
                            {
                                keyFound = false;
                                break;
                            }
                            DataFormatDataRow target =
                                idsItems.First(p => p.ActuallineNumber == idsItems.Max(q => q.ActuallineNumber));
                            try
                            {
                                if (!double.TryParse(target.Measured, out double tmp))
                                {
                                    continue;
                                }

                                valueList.Add(double.Parse(DataConvertor.ConvertUnits(target.Measured + target.Unit)));
                                limitList.Add(double.Parse(DataConvertor.ConvertUnits(target.High + target.Unit)));
                            }
                            catch (Exception)
                            {
                                EfuseStatic.Result = EfuseCheckResultType.Exception;
                            }
                        }
                        if (keyFound)
                        {
                            excelWorksheet.Cells[rowindex, colIndex + 1].Value = valueList.Sum();
                        }

                        idsLimits.TryAdd(setWriteItem.Key, limitList);
                    }
                    else //hardip
                    {
                        if (setWriteItem.Key.EqualsIgnoreCase("CONFIG#ane_adclk_pll0_fcal"))
                        {
                        }
                        var hipItems =
                            setHipDicItems.Where(
                                p =>
                                    p.Key.EqualsIgnoreCase(setWriteItem.Value.ReferenceKey) &&
                                    p.Site == setWriteItem.Value.Site
                                    && p.Number < setWriteItem.Value.Number).ToList();
                        if (setWriteItem.Value.ReferenceKey == "BTS_TS004_Temp10")
                        {
                        }
                        if (hipItems.Count > 0)
                        {
                            bool isNeedReverse = !string.IsNullOrEmpty(setWriteItem.Value.ReverseBit);
                            excelWorksheet.Cells[rowindex, colIndex + 1].Value =
                                isNeedReverse
                                    ? Math.Round(double.Parse(setWriteItem.Value.ReferenceValue))
                                    : Math.Round(double.Parse(hipItems.First(q => q.Number == hipItems.Max(p => p.Number)).Value));
                        }
                    }
                    bool compareFlag = false;
                    if (!double.TryParse(bdfItem.Resolution, out double resolution))
                    {
                        compareFlag = excelWorksheet.Cells[rowindex, colIndex].Text ==
                                           excelWorksheet.Cells[rowindex, colIndex + 1].Text;
                    }
                    else
                    {
                        try
                        {
                            resolution /= 1000;
                            double meas = Math.Round(double.Parse(excelWorksheet.Cells[rowindex, colIndex + 1].Text) * 1000) / 1000;
                            double fuseCal = double.Parse(excelWorksheet.Cells[rowindex, colIndex].Text) * resolution;
                            excelWorksheet.Cells[rowindex, colIndex].Value = fuseCal;
                            if (meas >= Math.Round((fuseCal - resolution) * 1000) / 1000 &&
                                meas <= Math.Round((fuseCal + resolution) * 1000) / 1000)
                            {
                                compareFlag = true;
                            }
                        }
                        catch (Exception)
                        {
                            EfuseStatic.Result = EfuseCheckResultType.Exception;
                        }
                    }
                    if (!compareFlag)
                    {
                        excelWorksheet.Cells[rowindex, colIndex].Value += "(F)";
                        excelWorksheet.Cells[rowindex, colIndex + 1].Value += "(F)";

                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_19, EnumErrorLevel.Error, excelWorksheet.Name, rowindex, colIndex, "Compared item not match.");
                        AddComment(appendRichText, excelWorksheet, error);
                    }
                    rowindex++;
                }
            }

            return rowindex;
        }

        public static void WriteEccCheckSheet(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, EfuseBitDefTable efuseBitDefTable, ref int eccDataStartCol)
        {
            if (excelWorksheet == null)
            {
                return;
            }

            int dataStartRow = 3;
            foreach (KeyValuePair<string, EccData> ecc in xDiceInfo.EccInfo)
            {
                string bank = WriteDsscNvSuffixRegex().Replace(ecc.Key, "").Replace("_", "");
                string accessMode = efuseBitDefTable.Rows.FirstOrDefault(x => x[efuseBitDefTable.BlockIdx].Replace("_", "").EqualsIgnoreCase(bank))![efuseBitDefTable.AccessModeIdx];
                string alg = accessMode.Split(':')[1].TrimEnd(')');
                excelWorksheet.Cells[dataStartRow - 2, 1].Value = bank.ToUpper();
                excelWorksheet.Cells[dataStartRow - 2, 1, dataStartRow - 2 + 1, 1].Merge = true;
                excelWorksheet.Cells[dataStartRow - 2, 1, dataStartRow - 2 + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[dataStartRow - 2, 1, dataStartRow - 2 + 1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelWorksheet.Cells[dataStartRow - 2, eccDataStartCol].Value = "Site" + xDiceInfo.Site;
                excelWorksheet.Cells[dataStartRow - 2, eccDataStartCol, dataStartRow - 2, eccDataStartCol + 1].Merge = true;
                excelWorksheet.Cells[dataStartRow - 2, eccDataStartCol, dataStartRow - 2, eccDataStartCol + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[dataStartRow - 1, eccDataStartCol].Value = "ECC";
                excelWorksheet.Cells[dataStartRow - 1, eccDataStartCol + 1].Value = "Data";

                int rowIdx = 0;
                foreach (string eccData in ecc.Value.RawDataList)
                {
                    (string dataLogRawEccData, string dataLogRawEccValue) = EfuseScriptUtility.GetEccAndData(ecc.Value.IsMsbFirst, alg, eccData);
                    string calcEccValue = EfuseScriptUtility.CalcEccValue(ecc.Value.IsMsbFirst, alg, dataLogRawEccData);
                    bool isMatch = dataLogRawEccValue.EqualsIgnoreCase(calcEccValue);

                    excelWorksheet.Cells[dataStartRow, 1].Value = "Row = " + rowIdx;
                    excelWorksheet.Cells[dataStartRow, eccDataStartCol].Value = dataLogRawEccValue;
                    excelWorksheet.Cells[dataStartRow, eccDataStartCol + 1].Value = dataLogRawEccData;
                    dataStartRow++;
                    rowIdx++;

                    if (!isMatch)
                    {
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_23, EnumErrorLevel.Error, "", dataStartRow, eccDataStartCol, calcEccValue);
                        Writer.HighLightCell(excelWorksheet, dataStartRow, eccDataStartCol, dataStartRow, eccDataStartCol, Color.Red);
                        AddComment(appendRichText, excelWorksheet, error);
                    }
                }

                dataStartRow += 3;
            }

            eccDataStartCol += 2;
        }

        public static void PrintError(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, Error error, Color color)
        {
            Writer.HighLightCell(excelWorksheet, error.RowNum, error.ColNum, error.RowNum, error.ColNum, color);
            if (!string.IsNullOrEmpty(error.Message))
            {
                AddComment(appendRichText, excelWorksheet, error);
            }
        }

        public static void PrintErrorComment(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, Error error)
        {
            excelWorksheet.Cells[error.RowNum, error.ColNum].Value += "(F)";
            if (!string.IsNullOrEmpty(error.Message))
            {
                AddComment(appendRichText, excelWorksheet, error);
            }
        }

        public static void PrintWarningComment(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, Error error)
        {
            excelWorksheet.Cells[error.RowNum, error.ColNum].Value += "(W)";
            if (!string.IsNullOrEmpty(error.Message))
            {
                AddComment(appendRichText, excelWorksheet, error);
            }
        }

        public static void AddComment(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, Error error)
        {
            try
            {
                error.SheetName = excelWorksheet.Name;
                if (error.ErrorLevel == EnumErrorLevel.Error)
                {
                    Writer.HighLightCell(excelWorksheet, error.RowNum, error.ColNum, error.RowNum, error.ColNum, Color.Red);
                }
                else if (error.ErrorLevel == EnumErrorLevel.Warning)
                {
                    Writer.HighLightCell(excelWorksheet, error.RowNum, error.ColNum, error.RowNum, error.ColNum, Color.LightPink);
                }
                excelWorksheet.Comments.Add(excelWorksheet.Cells[error.RowNum, error.ColNum], error.Message, "TAutoGen");
                if (!EFuseAppExcelWriterStatic.SummaryErrorsList.Exists(x => x.SheetName == excelWorksheet.Name && x.RowNum.Equals(error.RowNum) && x.ColNum.Equals(error.ColNum)))
                {
                    EFuseAppExcelWriterStatic.SummaryErrorsList.Add(error);
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

        //print BDF data
        public static void PrintBdfContent(List<XParseDatalog.EFuseSyntaxChkItem> eFuseSyntaxChkItems, int testCat, Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, EfuseBitDefTable efuseBitDefTable, List<string> oneRow, ref int lastAdRow, EfuseScriptConfig efuseScriptConfig)
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

            if (Writer.IsNeedHighlightRow(oneRow[efuseBitDefTable.PrgStageIdx], efuseScriptConfig))
            {
                Writer.HighLightCell(excelWorksheet, lastAdRow, 1, lastAdRow, efuseBitDefTable.DefaultValueIdx, Color.Yellow);
            }

            //if Syntax check items not exist in datalog, skip them
            if (eFuseSyntaxChkItems == null)
            {
                return;
            }

            if (oneRow[efuseBitDefTable.DefaultOrRealIdx].EqualsIgnoreCase("real"))
            {
                Writer.HighLightCell(excelWorksheet, lastAdRow, efuseBitDefTable.DefaultOrRealIdx + 1, lastAdRow, efuseBitDefTable.DefaultOrRealIdx + 1, Color.PaleGoldenrod);
                XParseDatalog.EFuseSyntaxChkItem? syntaxitem = eFuseSyntaxChkItems.FirstOrDefault(p => p.Id.EqualsIgnoreCase(oneRow[efuseBitDefTable.NameIdx]));
                if (syntaxitem != null && EfuseScriptUtility.GetJobOrder(oneRow[efuseBitDefTable.PrgStageIdx]) <= testCat)
                {
                    if (!EfuseScriptUtility.CompareValue(appendRichText, oneRow[efuseBitDefTable.LowLimitIdx], syntaxitem.LowLimit))
                    {
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_22, EnumErrorLevel.Error, "", lastAdRow, efuseBitDefTable.LowLimitIdx + 1, syntaxitem.LowLimit);
                        PrintError(appendRichText, excelWorksheet, error, Color.Yellow);
                    }
                    if (!EfuseScriptUtility.CompareValue(appendRichText, oneRow[efuseBitDefTable.HighLimitIdx], syntaxitem.HighLimit))
                    {

                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_21, EnumErrorLevel.Error, "", lastAdRow, efuseBitDefTable.HighLimitIdx + 1, syntaxitem.HighLimit);
                        PrintError(appendRichText, excelWorksheet, error, Color.Yellow);
                    }
                }
            }
            // check the shadow fields
            if (oneRow[efuseBitDefTable.NameIdx].EndsWithIgnoreCase("_shadow"))
            {
                string nonShadowName = oneRow[efuseBitDefTable.NameIdx].ToLower().Replace("_shadow", "");
                List<string>? nonShadowField = efuseBitDefTable.Rows.FirstOrDefault(x => x[efuseBitDefTable.NameIdx].EqualsIgnoreCase(nonShadowName));

                if (nonShadowField != null)
                {
                    if (oneRow[efuseBitDefTable.PrgStageIdx].EqualsIgnoreCase(nonShadowField[efuseBitDefTable.PrgStageIdx]))
                    {
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_28, EnumErrorLevel.Warning, "", lastAdRow, efuseBitDefTable.NameIdx + 1, $"The programming stage is same between {oneRow[efuseBitDefTable.NameIdx]} & {nonShadowField[efuseBitDefTable.NameIdx]}");
                        PrintError(appendRichText, excelWorksheet, error, Color.Yellow);
                    }
                }
            }
        }

        public static void CheckIedaNull(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, int colIndex)
        {
            for (int row = 6; row < 12; row++)
            {
                if (excelWorksheet.Cells[row, colIndex].Value == null)
                {
                    var error = new Error(EfuseCheckCmdLibError.E_MissingValue_02, EnumErrorLevel.Error, excelWorksheet.Name, row, colIndex, EfuseCheckCmdLibError.E_MissingValue_02.MessageTemplate);
                    Writer.HighLightCell(excelWorksheet, row, colIndex, row, colIndex, Color.Red);
                    AddComment(appendRichText, excelWorksheet, error);
                }
            }
        }

        public static void CheckPrr(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, int colIndex, bool isNeedReversed = false)
        {
            if (string.IsNullOrEmpty(xDiceInfo.PrrCode))
            {
                Writer.HighLightCell(excelWorksheet, 5, colIndex, 5, colIndex, Color.Yellow);
            }
            else
            {
                string calPrr = EfuseScriptUtility.GetProberHexcode(xDiceInfo, isNeedReversed);
                var prrError = new Error(EfuseCheckCmdLibError.E_MismatchValue_29, EnumErrorLevel.Error, excelWorksheet.Name, 5, colIndex, $"{xDiceInfo.PrrLotId}_W{xDiceInfo.PrrWaferId}_X{xDiceInfo.PrrX}_Y{xDiceInfo.PrrY},Calculate_PRR:{calPrr}");
                if (xDiceInfo.PrrCode != calPrr)
                {
                    PrintError(appendRichText, excelWorksheet, prrError, Color.Red);
                    EFuseAppExcelWriterStatic.GeneralCheckerErrors.Add(prrError);
                }
                else
                {
                    excelWorksheet.Comments.Add(excelWorksheet.Cells[prrError.RowNum, prrError.ColNum], prrError.Message, "TAutoGen");
                }
            }
        }

        public static void CheckHramEcid(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, int colIndex)
        {
            string reverseHramEcid = "";
            if (!string.IsNullOrEmpty(xDiceInfo.HramEcid53Bit))
            {
                IEnumerable<char> reverse = xDiceInfo.HramEcid53Bit.Reverse();
                foreach (char val in reverse)
                {
                    reverseHramEcid += val;
                }
                string hexstr = Convert.ToInt64(reverseHramEcid, 2).ToString("X").PadLeft(xDiceInfo.PrrCode.Length, '0');

                if (xDiceInfo.PrrCode != hexstr)
                {
                    var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_30, EnumErrorLevel.Error, excelWorksheet.Name, 10, colIndex, "Not Match with PRR_Code.");
                    Writer.HighLightCell(excelWorksheet, 10, colIndex, 10, colIndex, Color.Red);
                    AddComment(appendRichText, excelWorksheet, error);
                }
            }
        }

        public static void CheckSvmcFuse(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, XParseDatalog.StageDicesInfo stageDicesInfo, int colIndex, EfuseCfgTable efuseCfgTable)
        {
            if (!string.IsNullOrEmpty(xDiceInfo.SvmCFuse288Bits))
            {
                var fuseRevisionRow = xDiceInfo.AllReadFromDssc.Where(p => p.Id.EqualsIgnoreCase("fuse_revision")).ToList();
                bool isNeedReversed = fuseRevisionRow.FirstOrDefault()!.Order.StartsWith("(MSB)") || fuseRevisionRow.First().Block.EqualsIgnoreCase("config");
                string fuseRevisionRawData = isNeedReversed ? fuseRevisionRow.Select(p => p.RawData).ToList()[0] : new string([.. fuseRevisionRow.Select(p => p.RawData).ToList()[0].Reverse()]);

                // Check for CP, WLFT stage
                if (CpRegex().IsMatch(stageDicesInfo.ChannelMapStage) || WlftRegex().IsMatch(stageDicesInfo.ChannelMapStage))
                {
                    if (xDiceInfo.SvmCFuse288Bits != fuseRevisionRawData)
                    {
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_24, EnumErrorLevel.Error, excelWorksheet.Name, 11, colIndex, EfuseCheckCmdLibError.E_MismatchValue_24.MessageTemplate);
                        Writer.HighLightCell(excelWorksheet, 11, colIndex, 11, colIndex, Color.Red);
                        AddComment(appendRichText, excelWorksheet, error);
                    }
                }
                // Check for FT, SLT stage
                else
                {
                    var cfgScenarios = new List<string>();
                    int cfgScenarioIndx = efuseCfgTable.Scenario[XParseDatalog.ScenarioInDatalog];
                    string stage = "";
                    string cfgScenario = "";
                    int bitWidth = 0;

                    foreach (List<string> rows in efuseCfgTable.CfgRows)
                    {
                        if (rows.Count < cfgScenarioIndx)
                        {
                            continue;
                        }

                        stage = rows[efuseCfgTable.PrgStageIdx];
                        cfgScenario = rows[cfgScenarioIndx];
                        bitWidth = (int)Convert.ToInt64(rows[efuseCfgTable.BitWidthIdx]);

                        if (EfuseScriptUtility.GetJobOrder(stage) <= EfuseScriptUtility.GetJobOrder(XParseDatalog.CurrentJob))
                        {
                            string binaryCfg = Convert.ToString(Convert.ToInt32(cfgScenario, 16), 2).PadLeft(bitWidth, '0');
                            cfgScenarios.Add(binaryCfg);
                        }
                        else
                        {
                            string cfg = "";
                            cfg = cfg.PadLeft(bitWidth, '0');
                            cfgScenarios.Add(cfg);
                        }
                    }

                    string result = "";
                    foreach (string str in cfgScenarios)
                    {
                        char[] chars = str.ToCharArray();
                        Array.Reverse(chars);
                        string strReverse = new(chars);
                        result += strReverse;
                    }

                    if (xDiceInfo.SvmCFuse288Bits != result)
                    {
                        var error = new Error(EfuseCheckCmdLibError.E_MismatchValue_24, EnumErrorLevel.Error, excelWorksheet.Name, 11, colIndex, EfuseCheckCmdLibError.E_MismatchValue_24.MessageTemplate);
                        Writer.HighLightCell(excelWorksheet, 11, colIndex, 11, colIndex, Color.Red);
                        AddComment(appendRichText, excelWorksheet, error);
                    }
                }
            }
        }

        public static void CheckSvmcFuseSiteFail(Action<string, string> appendRichText, ExcelWorksheet excelWorksheet, XDiceInfo xDiceInfo, int colIndex)
        {
            foreach (EfuseDatalogItem onePair in xDiceInfo.AllReadFromDssc)
            {
                XBitDefRow? row = XReadEfuseBitDef.GetRow(onePair.Id, onePair.Block);
                if (row != null)
                {
                    string name = row.Name.Replace(row.Block + "#", "");
                    //if (row != null && Regex.IsMatch(row.name, @".*FUSE_REVISION", RegexOptions.IgnoreCase))
                    if (row != null && FuseRevisionExactRegex().IsMatch(name))
                    {
                        if (!EfuseScriptUtility.CheckEfuseDefaultValue(appendRichText, onePair.Value, row.DefaultValue))
                        {
                            var error = new Error(EfuseCheckCmdLibError.W_MismatchValue_02, EnumErrorLevel.Warning, excelWorksheet.Name, 11, colIndex, EfuseCheckCmdLibError.W_MismatchValue_02.MessageTemplate);
                            Writer.HighLightCell(excelWorksheet, 11, colIndex, 11, colIndex, Color.Red);
                            AddComment(appendRichText, excelWorksheet, error);
                        }
                    }
                }
            }
        }

        public static void FreezeSheetFormat(ExcelWorkbook excelWorkbook, string stage)
        {
            foreach (ExcelWorksheet ws in excelWorkbook.Worksheets)
            {
                string locAddr = ws.Dimension.Address;
                IExcelConditionalFormattingContainsText cf4 = ws.Cells[locAddr].ConditionalFormatting.AddContainsText();
                cf4.Text = "(F)";
                cf4.Style.Fill.BackgroundColor.Color = Color.Red;
                cf4.Style.Font.Bold = true;
                if (ws.Name.StartsWith(SheetConst.Type5BitDefTable))
                {
                    ws.Column(1).Width = 40;
                    //ws.View.FreezePanes(7, 19);
                    ws.View.FreezePanes(13, 19);
                    IExcelConditionalFormattingContainsText cf = ws.Cells[locAddr].ConditionalFormatting.AddContainsText();
                    cf.Text = stage;
                    cf.Style.Fill.BackgroundColor.Color = Color.LightGray;
                    cf.Style.Font.Bold = true;
                }
                else if (ws.Name == HipbvRealValueSheetName || ws.Name == SetWriteVariableSheetName)
                {
                    ws.View.FreezePanes(3, 38);
                }
                else if (ws.Name == FusePattenSummarySheetName)
                {
                    ws.View.FreezePanes(2, ws.Dimension.Columns);
                    ws.Cells.TryAutoFitColumns();
                }
                else if (ws.Name == SummaryReportSheetName)
                {
                    ws.View.FreezePanes(10, 6);
                    ws.Cells.TryAutoFitColumns();
                    ws.Cells[9, 1, 9, 6].AutoFilter = true;
                }
                else if (ws.Name == EccCheckSheetName)
                {
                    ws.Cells.TryAutoFitColumns();
                }
                else
                {
                    ws.View.FreezePanes(5, 7);
                }
            }
        }
    }
}
