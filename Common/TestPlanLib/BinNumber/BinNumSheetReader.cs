using System.Collections.Generic;
using System.Linq;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

namespace TestPlanLib.BinNumber
{
    public class BinNumSheetReader
    {
        private const string ConModule = "Module";
        private const string ConCategory1 = "Category1";
        private const string ConCategory2 = "Category2";
        private const string ConBinFirst = "BinFirst";
        private const string ConBinLast = "BinLast";
        private const string ConHardBin = "HardBin";
        private const string ConResult = "Result";
        private const string ConComment = "Comment";

        private int _startRow;
        private int _columnModule;
        private int _columnCategory1;
        private int _columnCategory2;
        private int _columnConBinFirst;
        private int _columnConBinLast;
        private int _columnHardBin;
        private int _columnResult;
        private int _columnComment;

        private ExcelWorksheet? _worksheet;
        public HashSet<BinNumInfo> ReadSheet(ExcelWorksheet excelWorksheet)
        {
            _worksheet = excelWorksheet;
            HashSet<BinNumInfo> binNumDefSet = [];

            ReadHeader();

            ReadData(binNumDefSet);

            return binNumDefSet;
        }

        private void ReadData(HashSet<BinNumInfo> binNumInfoCollect)
        {
            for (int i = _startRow + 1; i <= _worksheet!.Dimension.End.Row; i++)
            {
                if (string.IsNullOrEmpty(EpplusExtensions.GetCellValue(_worksheet!, i, _columnModule).Trim()))
                {
                    continue;
                }

                if (!int.TryParse(EpplusExtensions.GetCellValue(_worksheet!, i, _columnConBinFirst).Trim(), out int binFirst))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_23, _worksheet!.Name, i, _columnConBinFirst, $"Can't parse {ConBinFirst} in {_worksheet!.Name}:Row{i}, it should be a integer.", [ConBinFirst, _worksheet!.Name, i.ToString()]);
                }
                if (!int.TryParse(EpplusExtensions.GetCellValue(_worksheet!, i, _columnConBinLast).Trim(), out int binLast))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_24, _worksheet!.Name, i, _columnConBinLast, $"Can't parse {ConBinLast} in {_worksheet!.Name}:Row{i}, it should be a integer.", [ConBinLast, _worksheet!.Name, i.ToString()]);
                }
                if (!int.TryParse(EpplusExtensions.GetCellValue(_worksheet!, i, _columnHardBin).Trim(), out int hardBin))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_25, _worksheet!.Name, i, _columnHardBin, $"Can't parse {ConHardBin} in {_worksheet!.Name}:Row{i}, it should be a integer.", [ConHardBin, _worksheet!.Name, i.ToString()]);
                }

                if (binFirst == 0 && binLast != 0)
                {
                    binFirst = binLast;
                }
                else if (binLast == 0 && binFirst != 0)
                {
                    binLast = binFirst;
                }

                if (binFirst > binLast)
                {
                    (binFirst, binLast) = (binLast, binFirst);
                }
                if (binFirst < 1)
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_26, _worksheet!.Name, i, i, $"{ConBinFirst}:{binFirst} in {_worksheet!.Name}:Row{i}, should be in the range of 1-9999.", [ConBinFirst, binFirst.ToString(), _worksheet!.Name, i.ToString()]);
                    binFirst = 1;
                }
                if (binLast > 9999)
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_27, _worksheet!.Name, i, i, $"{ConBinLast}:{binLast} in {_worksheet!.Name}:Row{i}, should be in the range of 1-9999.", [ConBinLast, binLast.ToString(), _worksheet!.Name, i.ToString()]);
                    binLast = 9999;
                }

                string module = EpplusExtensions.GetCellValue(_worksheet!, i, _columnModule).Replace("_", "").Trim();
                string category1 = EpplusExtensions.GetCellValue(_worksheet!, i, _columnCategory1).Replace("_", "").Trim();
                string category2 = EpplusExtensions.GetCellValue(_worksheet!, i, _columnCategory2).Replace("_", "").Trim();
                category1 = category1.EqualsIgnoreCase("X") ? "" : category1;
                category2 = category2.EqualsIgnoreCase("X") ? "" : category2;

                BinNumInfo binNumInfo = new BinNumInfo
                {
                    Module = module,
                    Category1 = category1,
                    Category2 = category2,
                };

                if (!binNumInfoCollect.Add(binNumInfo))
                {
                    binNumInfo = binNumInfoCollect.FirstOrDefault(x => x.Equals(binNumInfo))!;
                }

                BinNumDef binNumDef = new BinNumDef
                {
                    BinFirst = binFirst,
                    BinLast = binLast,
                    HardBin = hardBin,
                    Result = EpplusExtensions.GetCellValue(_worksheet!, i, _columnResult).Trim(),
                    Comment = EpplusExtensions.GetCellValue(_worksheet!, i, _columnComment).Trim(),
                    BinNumSheetRowNum = i
                };
                binNumInfo.BinNumDef.Add(binNumDef);
            }
        }

        private bool ReadHeader()
        {
            for (int i = 1; i <= _worksheet!.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= _worksheet!.Dimension.End.Column; j++)
                {
                    if (EpplusExtensions.GetCellValue(_worksheet!, i, j).Replace(" ", "").EqualsIgnoreCase(ConModule))
                    {
                        _startRow = i;
                        _columnModule = j;
                        break;
                    }
                }
                if (_startRow != 0)
                {
                    break;
                }
            }

            if (_startRow == 0 || _columnModule == 0)
            {
                string msg = $"Can't get bin number sheet in test plan, missing {ConModule} header, " + $"it should have {ConModule}, {ConCategory1}, {ConCategory2}, {ConBinFirst}, {ConBinLast}, {ConHardBin}, {ConResult}, {ConComment}";
                Response.Report(msg, EnumMessageLevel.Error);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_28, _worksheet!.Name, 0, 0, $"Can't get bin number sheet in test plan, missing {ConModule} header, " + $"it should have {ConModule}, {ConCategory1}, {ConCategory2}, {ConBinFirst}, {ConBinLast}, {ConHardBin}, {ConResult}, {ConComment}", [ConModule, ConCategory1, ConCategory2, ConBinFirst, ConBinLast, ConHardBin, ConResult, ConComment]);
                return false;
            }

            for (int i = _columnModule; i <= _worksheet!.Dimension.End.Column; i++)
            {
                GetIndex(i);
            }

            string missingHeaderMsg = GetMissingHeaderMsg();
            if (missingHeaderMsg.Length != 0)
            {
                string msg = $"Can't get bin number sheet in test plan, missing {missingHeaderMsg} header.";
                Response.Report(msg, EnumMessageLevel.Error);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_29, _worksheet!.Name, _startRow, _startRow, $"Can't get bin number sheet in test plan, missing {missingHeaderMsg} header.", [missingHeaderMsg]);
                return false;
            }
            return true;
        }

        private string GetMissingHeaderMsg()
        {
            string missingHeaderMsg = "";

            if (_columnCategory1 == 0)
            {
                missingHeaderMsg += $", {ConCategory1}";
            }
            if (_columnCategory2 == 0)
            {
                missingHeaderMsg += $", {ConCategory2}";
            }
            if (_columnConBinFirst == 0)
            {
                missingHeaderMsg += $", {ConBinFirst}";
            }
            if (_columnConBinLast == 0)
            {
                missingHeaderMsg += $", {ConBinLast}";
            }
            if (_columnHardBin == 0)
            {
                missingHeaderMsg += $", {ConHardBin}";
            }
            if (_columnResult == 0)
            {
                missingHeaderMsg += $", {ConResult}";
            }
            if (_columnComment == 0)
            {
                missingHeaderMsg += $", {ConComment}";
            }

            return missingHeaderMsg;
        }

        private void GetIndex(int i)
        {
            if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConCategory1) && _columnCategory1 == 0)
            {
                _columnCategory1 = i;
            }
            else if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConCategory2) && _columnCategory2 == 0)
            {
                _columnCategory2 = i;
            }
            else if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConBinFirst) && _columnConBinFirst == 0)
            {
                _columnConBinFirst = i;
            }
            else if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConBinLast) && _columnConBinLast == 0)
            {
                _columnConBinLast = i;
            }
            else if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConHardBin) && _columnHardBin == 0)
            {
                _columnHardBin = i;
            }
            else if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConResult) && _columnResult == 0)
            {
                _columnResult = i;
            }
            else if (EpplusExtensions.GetCellValue(_worksheet!, _startRow, i).Replace(" ", "").EqualsIgnoreCase(ConComment) && _columnComment == 0)
            {
                _columnComment = i;
            }
        }
    }
}
