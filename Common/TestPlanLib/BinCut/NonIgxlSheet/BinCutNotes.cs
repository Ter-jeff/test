using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet
{
    public partial class BinCutNotes(ExcelWorksheet excelWorksheet, string folder = "") : BinCutNonIgxlBase(excelWorksheet, folder)
    {
        #region Field
        private const string ConHeaderBaseVoltage = "Base Voltage (mV)*";

        [GeneratedRegex(ConHeaderBaseVoltage, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public ExcelWorksheet NotesSheet = excelWorksheet;

        #endregion

        public double FindStepSize()
        {
            DataTable dt = ReadSheet(NotesSheet);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString()!.Contains("a multiple of Step"))
                    {
                        string size = dt.Rows[i][j].ToString()!.Split('=').Last().Split(',').First().Replace("mV", "");
                        if (double.TryParse(size, out double stepSzie))
                        {
                            StepSize = stepSzie;
                            return stepSzie;
                        }
                        else
                        {
                            return 0;
                        }
                    }
                }
            }
            return 0;
        }

        public double FindBaseVoltage()
        {
            DataTable dt = ReadSheet(NotesSheet);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (_regex.IsMatch(dt.Rows[j][i].ToString()!))
                    {
                        if (double.TryParse(dt.Rows[j][i + 1].ToString(), out double baseVolate))
                        {
                            BaseVoltage = baseVolate;
                            return baseVolate;
                        }
                        return 0;
                    }
                }
            }
            return 0;
        }
    }
}
