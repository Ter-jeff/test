using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.DataStruct
{
    public class PowerInfoReader : MySheetReader<PowerInfoSheet>
    {
        private const string ConPowerInfoBodySymbol = "PinName";
        private const string ConPowerInfoChipletSymbol = "Chiplet";

        private Dictionary<string, int> _titleDic = [];

        public override PowerInfoSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            GetDimensions();

            PowerInfoSheet result = new PowerInfoSheet();
            GetFirstHeaderPosition([ConPowerInfoBodySymbol, ConPowerInfoChipletSymbol]);
            _titleDic = GetTittle();
            int pinNameCol = FindTitleIndex(_titleDic, ConPowerInfoBodySymbol);
            int pwrSeqCol = FindTitleIndex(_titleDic, "PowerUpSequence");
            if (pwrSeqCol == 0)
            {
                pwrSeqCol = FindTitleIndex(_titleDic, "PowerSequence");
            }

            Dictionary<int, string> iFoldCols = FindAllTitleIndex(_titleDic, "ifold");
            result.ExistEvs = iFoldCols.Any(x => x.Value.Contains("EVS", System.StringComparison.OrdinalIgnoreCase));
            result.ExistConti = iFoldCols.Any(x => x.Value.Contains("Conti", System.StringComparison.OrdinalIgnoreCase));
            result.ExistHip = iFoldCols.Any(x => x.Value.Contains("HIP", System.StringComparison.OrdinalIgnoreCase));
            int pwrDownSeqCol = FindTitleIndex(_titleDic, "PowerDownSequence");
            int chipletCol = FindTitleIndex(_titleDic, "Chiplet");
            int allSeq = 0;
            int pwrseq;
            for (int row = StartRow + 1; row <= ExcelWorksheet.Dimension.End.Row; row++)
            {
                bool success = int.TryParse(EpplusExtensions.GetCellValue(ExcelWorksheet, row, pwrSeqCol), out pwrseq);
                if (success && pwrseq > allSeq)
                {
                    allSeq = pwrseq;
                }
            }
            for (int row = StartRow + 1; row <= ExcelWorksheet.Dimension.End.Row; row++)
            {
                PowerInfoRow info = new PowerInfoRow
                {
                    Ifold = AnalysisIfoldByStage(row, iFoldCols),
                    PinName = EpplusExtensions.GetCellValue(ExcelWorksheet, row, pinNameCol)
                };
                if (info.PinName?.Length == 0)
                {
                    break;
                }
                info.PowerSequence = EpplusExtensions.GetCellValue(ExcelWorksheet, row, pwrSeqCol);
                info.Chiplet = EpplusExtensions.GetCellValue(ExcelWorksheet, row, StartCol);
                info.RowNum = row;
                //info.PowerDownSequence = pwrDownSeqCol == 0 ? info.PowerSequence : ExcelReader.GetCellValue(_currentSheet, row, pwrDownSeqCol);
                if (pwrDownSeqCol != 0)
                {
                    info.PowerDownSequence = EpplusExtensions.GetCellValue(ExcelWorksheet, row, pwrDownSeqCol);
                }
                else
                {
                    bool success = int.TryParse(EpplusExtensions.GetCellValue(ExcelWorksheet, row, pwrSeqCol), out pwrseq);
                    if (success)
                    {
                        info.PowerDownSequence = (1 + allSeq - pwrseq).ToString();
                    }
                    else
                    {
                        info.PowerDownSequence = info.PowerSequence;
                    }
                }
                if (chipletCol != 0)
                {
                    info.Chiplet = EpplusExtensions.GetCellValue(ExcelWorksheet, row, chipletCol);
                }

                result.Rows.Add(info);
            }

            return result;
        }

        private Dictionary<string, int> GetTittle()
        {
            Dictionary<string, int> result = [];
            for (int i = 1; i <= ExcelWorksheet.Dimension.End.Column; i++)
            {
                if (EpplusExtensions.GetCellValue(ExcelWorksheet, StartRow, i)?.Length == 0)
                {
                    break;
                }

                result.Add(EpplusExtensions.GetCellValue(ExcelWorksheet, StartRow, i), i);
            }
            return result;
        }

        private static int FindTitleIndex(Dictionary<string, int> title, string titleName)
        {
            Regex titleReg = new Regex(titleName, RegexOptions.IgnoreCase);
            KeyValuePair<string, int> entry = title.FirstOrDefault(x => titleReg.IsMatch(x.Key));
            return entry.Key != null ? entry.Value : 0;
        }

        private static Dictionary<int, string> FindAllTitleIndex(Dictionary<string, int> title, string titleName)
        {
            var result = new Dictionary<int, string>();
            Regex titleReg = new Regex(titleName, RegexOptions.IgnoreCase);
            List<string> entries = title.Keys.ToList().FindAll(titleReg.IsMatch);
            if (entries.Count == 0)
            {
                result.Add(0, "");
            }

            foreach (string entry in entries)
            {
                string job = entry.Contains(':') ? entry.Split(':').Last() : "";
                result.Add(title[entry], job);
            }
            return result;
        }

        private string AnalysisIfoldByStage(int row, Dictionary<int, string> cols)
        {
            var tmpList = new List<string>();
            foreach (KeyValuePair<int, string> col in cols)
            {
                string value = col.Key == 0 ? "" : EpplusExtensions.GetCellValue(ExcelWorksheet, row, col.Key);
                if (string.IsNullOrEmpty(col.Value))
                {
                    tmpList.Add(value);
                }
                else
                {
                    tmpList.Add($"{value}:{col.Value}");
                }
            }
            return string.Join(";", tmpList);
        }
    }
}
