using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.SelSram
{
    public class InputSelSram
    {
        private int _combosWritePos = -1;
        private int _combosReadPos = -1;
        private int _combosCheckPos = -1;
        private int _combosStartRowPos = -1;
        private int _patOrgPos = -1;
        private int _patDigcapPos = -1;
        private int _patStartRowPos = -1;
        private int _bitsLogicPinsPos = -1;
        private int _bitsStartRowPos = -1;
        internal int _otherPatternsPos;
        internal int _selSramBitsPos = -1;
        internal int _selSramDigcapPos = -1;
        internal int _otherPatStartRowPos = -1;
        private readonly List<int> _patternsPos = new List<int>();

        private const string ConCombosWrite = "WRITE";
        private const string ConCombosRead = "READ";
        private const string ConCombosCheck = "CHECK";
        private const string ConOrg = "Org";
        private const string ConReadbackPat = "SELSRAM DigCap Readback Pattern";
        private const string ConBitsLogicPins = "Logic Pins";
        private const string ConOtherPatterns = "Other Patterns";
        private const string ConSelSramBits = "SELSRAM bits";
        private readonly Regex _regColHeaderPattern = new Regex(@"^Pattern\d*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public void ParsingReadbackSheet()
        {
            ExcelWorksheet sheetReadbackCombos = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Readback_Combos"];
            if (sheetReadbackCombos != null)
            {
                GetReadBackCombosHeaderPos(sheetReadbackCombos);
                AddingCombosData2List(sheetReadbackCombos);
            }

            ExcelWorksheet sheetReadbackPat = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Readback_Patterns"];
            if (sheetReadbackPat != null)
            {
                GetReadBackPatHeaderPos(sheetReadbackPat);
                AddingPatternData2Dic(sheetReadbackPat);
            }

            ExcelWorksheet sheetReadbackBits = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Readback_Bits"];
            if (sheetReadbackBits != null)
            {
                GetReadBackBitsHeaderPos(sheetReadbackBits);
                ProcessingReadBackBits(sheetReadbackBits);
            }

            ExcelWorksheet sheetReadbackOtherPatterns = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Readback_Other_Patterns"];
            if (sheetReadbackOtherPatterns != null)
            {
                GetReadbackOtherPatHeaderPos(sheetReadbackOtherPatterns);
                AddingOtherPat2Dic(sheetReadbackOtherPatterns);
            }
        }

        private void ProcessingReadBackBits(ExcelWorksheet sheetReadbackBits)
        {
            SelSramPatternSingleton selSram = SelSramPatternSingleton.GetInstance();
            for (int i = _bitsStartRowPos; i <= sheetReadbackBits.Dimension.End.Row; ++i)
            {
                string pinname = EpplusExtensions.GetCellValue(sheetReadbackBits, i, _bitsLogicPinsPos).ToUpper();
                selSram.Add2BitsLogicPins(pinname);
            }
        }

        private void GetReadBackBitsHeaderPos(ExcelWorksheet sheetReadbackBits)
        {
            for (int row = 1; row < 10; ++row)
            {
                for (int col = 1; col < 10; ++col)
                {
                    string header = EpplusExtensions.GetCellValue(sheetReadbackBits, row, col);
                    if (Regex.IsMatch(header, ConBitsLogicPins, RegexOptions.IgnoreCase))
                    {
                        _bitsLogicPinsPos = col;
                    }

                    if (_bitsLogicPinsPos != -1)
                    {
                        _bitsStartRowPos = row + 1;
                        return;
                    }
                }
            }
        }

        internal void GetReadbackOtherPatHeaderPos(ExcelWorksheet sheetReadbackOtherPatterns)
        {
            for (int row = 1; row < 10; ++row)
            {
                for (int col = 1; col < 10; ++col)
                {
                    string header = EpplusExtensions.GetCellValue(sheetReadbackOtherPatterns, row, col);
                    if (Regex.IsMatch(header, ConOtherPatterns, RegexOptions.IgnoreCase))
                    {
                        _otherPatternsPos = col;
                    }
                    else if (Regex.IsMatch(header, ConSelSramBits, RegexOptions.IgnoreCase))
                    {
                        _selSramBitsPos = col;
                    }
                    else if (Regex.IsMatch(header, ConReadbackPat, RegexOptions.IgnoreCase))
                    {
                        _selSramDigcapPos = col;
                    }

                    if (_otherPatternsPos != -1 && _selSramBitsPos != -1 && _selSramDigcapPos != -1)
                    {
                        _otherPatStartRowPos = row + 1;
                        return;
                    }
                }
            }
        }

        private void GetReadBackPatHeaderPos(ExcelWorksheet sheetReadbackPat)
        {
            for (int row = 1; row < 10; ++row)
            {
                for (int col = 1; col < 10; ++col)
                {
                    string header = EpplusExtensions.GetCellValue(sheetReadbackPat, row, col);
                    if (Regex.IsMatch(header, ConOrg, RegexOptions.IgnoreCase))
                    {
                        _patOrgPos = col;
                    }
                    else if (Regex.IsMatch(header, ConReadbackPat, RegexOptions.IgnoreCase))
                    {
                        _patDigcapPos = col;
                    }
                    else if (_regColHeaderPattern.IsMatch(header))
                    {
                        _patternsPos.Add(col);
                    }
                }
                if (_patOrgPos != -1 && _patDigcapPos != -1)
                {
                    _patStartRowPos = row + 1;
                    return;
                }
            }
        }

        internal void AddingOtherPat2Dic(ExcelWorksheet sheetReadbackOtherPatterns)
        {
            SelSramPatternSingleton selSram = SelSramPatternSingleton.GetInstance();
            for (int i = _otherPatStartRowPos; i <= sheetReadbackOtherPatterns.Dimension.End.Row; ++i)
            {
                string otherPat = EpplusExtensions.GetCellValue(sheetReadbackOtherPatterns, i, _otherPatternsPos).ToUpper();
                if (otherPat.Trim().Equals(""))
                {
                    break;
                }

                if (otherPat.Contains("*")) // HardIP_*
                {
                    foreach (ExcelWorksheet excelWorksheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        string sheetName = excelWorksheet.Name;
                        if (sheetName.ToUpper().Equals("HARDIP_DC"))
                        {
                            continue;
                        }

                        if (sheetName.StartsWith("HARDIP_", StringComparison.OrdinalIgnoreCase))
                        {
                            // rename from HardIP_DRAM_IDS to HardIP_DRAMIDS
                            sheetName = "HARDIP_" + sheetName.ToUpper().Replace("HARDIP_", "").Replace("_", "");
                            if (!selSram.DicOtherPat.ContainsKey(sheetName))
                            {
                                selSram.DicOtherPat.Add(sheetName, new OtherPatData());
                            }
                        }
                    }
                }
                else if (!selSram.DicOtherPat.ContainsKey(otherPat))
                {
                    selSram.DicOtherPat.Add(otherPat, new OtherPatData());
                }
            }
        }

        private void AddingPatternData2Dic(ExcelWorksheet sheetReadbackPat)
        {
            SelSramPatternSingleton selSram = SelSramPatternSingleton.GetInstance();
            for (int i = _patStartRowPos; i <= sheetReadbackPat.Dimension.End.Row; ++i)
            {
                string org = EpplusExtensions.GetCellValue(sheetReadbackPat, i, _patOrgPos).ToUpper();
                string pat = EpplusExtensions.GetCellValue(sheetReadbackPat, i, _patDigcapPos).ToUpper();
                if (org != "" && !selSram.DicReadbackPat.ContainsKey(org))
                {
                    selSram.DicReadbackPat.Add(org, pat.Equals("") ? "TBD" : pat);
                    if (selSram.DicInitPatterns.ContainsKey(org))
                    {
                        continue;
                    }
                    selSram.DicReadbackInitPat.Add(org, new List<string>());
                    foreach (int patternCol in _patternsPos)
                    {
                        string initPat = EpplusExtensions.GetCellValue(sheetReadbackPat, i, patternCol).ToUpper().Trim();
                        if (!string.IsNullOrEmpty(initPat))
                        {
                            selSram.DicReadbackInitPat[org].Add(initPat);
                        }
                    }
                }
            }
            selSram.HasReadBackInitPat = _patternsPos.Any();
        }

        private void AddingCombosData2List(ExcelWorksheet sheetReadbackCombos)
        {
            SelSramPatternSingleton selSram = SelSramPatternSingleton.GetInstance();
            for (int i = _combosStartRowPos; i <= sheetReadbackCombos.Dimension.End.Row; ++i)
            {
                selSram.AddCombosWrite(EpplusExtensions.GetCellValue(sheetReadbackCombos, i, _combosWritePos));
                selSram.AddCombosRead(EpplusExtensions.GetCellValue(sheetReadbackCombos, i, _combosReadPos));
                selSram.AddCombosCheck(EpplusExtensions.GetCellValue(sheetReadbackCombos, i, _combosCheckPos));
            }
        }

        private void GetReadBackCombosHeaderPos(ExcelWorksheet sheetReadbackCombos)
        {
            for (int row = 1; row < 10; ++row)
            {
                for (int col = 1; col < 10; ++col)
                {
                    string header = EpplusExtensions.GetCellValue(sheetReadbackCombos, row, col);
                    if (Regex.IsMatch(header, ConCombosWrite, RegexOptions.IgnoreCase))
                    {
                        _combosWritePos = col;
                    }
                    else if (Regex.IsMatch(header, ConCombosRead, RegexOptions.IgnoreCase))
                    {
                        _combosReadPos = col;
                    }
                    else if (Regex.IsMatch(header, ConCombosCheck, RegexOptions.IgnoreCase))
                    {
                        _combosCheckPos = col;
                    }

                    if (_combosWritePos != -1 && _combosReadPos != -1 && _combosCheckPos != -1)
                    {
                        _combosStartRowPos = row + 1;
                        return;
                    }
                }
            }
        }
    }
}
