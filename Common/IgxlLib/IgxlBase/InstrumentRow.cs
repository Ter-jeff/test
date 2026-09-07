using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using OfficeOpenXml;

namespace IgxlLib.IgxlBase
{
    public class InstrumentRow(string? min, string? max, string? currentType, string? isMerge, string? instrument, string? iFold)
    {
        public string? Min { get; set; } = min;
        public string? Max { get; set; } = max;
        public string? CurrentType { get; set; } = currentType;
        public string? IsMerge { get; set; } = isMerge;
        public string? Instrument { get; set; } = instrument;
        public string? IFold { get; set; } = iFold;

        public static List<InstrumentRow> Reader(string filename)
        {
            var listInstRow = new List<InstrumentRow>();
            if (!File.Exists(filename))
            {
                return listInstRow;
            }

            var dicHeader = new Dictionary<string, int>();

            using (var fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using var ep = new ExcelPackage(fs);
                ExcelWorksheet worksheet = ep.Workbook.Worksheets["IFold"];
                if (worksheet == null)
                {
                    return listInstRow;
                }

                // get header index
                for (int col = 1; col <= worksheet.Dimension.End.Column; ++col)
                {
                    string header = worksheet.Cells[1, col].Value != null ? worksheet.Cells[1, col].Value!.ToString()!.Trim() : "";
                    if (header.ContainsIgnoreCase("Min"))
                    {
                        dicHeader.Add("Min", col);
                    }
                    else if (header.ContainsIgnoreCase("Max"))
                    {
                        dicHeader.Add("Max", col);
                    }
                    else if (header.ContainsIgnoreCase("CurrentType"))
                    {
                        dicHeader.Add("CurrentType", col);
                    }
                    else if (header.ContainsIgnoreCase("IsMerge"))
                    {
                        dicHeader.Add("IsMerge", col);
                    }
                    else if (header.ContainsIgnoreCase("Instrument"))
                    {
                        dicHeader.Add("Instrument", col);
                    }
                    else if (header.ContainsIgnoreCase("Ifold"))
                    {
                        dicHeader.Add("Ifold", col);
                    }
                }
                if (dicHeader.Count < 6)
                {
                    return listInstRow;
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; ++row)
                {
                    string? min = worksheet.Cells[row, dicHeader["Min"]].Value?.ToString() ?? "";
                    string? max = worksheet.Cells[row, dicHeader["Max"]].Value?.ToString() ?? "";
                    string? currentType = worksheet.Cells[row, dicHeader["CurrentType"]].Value?.ToString() ?? "";
                    string? isMerge = worksheet.Cells[row, dicHeader["IsMerge"]].Value?.ToString() ?? "";
                    string? instrument = worksheet.Cells[row, dicHeader["Instrument"]].Value?.ToString() ?? "";
                    string? ifold = worksheet.Cells[row, dicHeader["Ifold"]].Value?.ToString() ?? "";
                    listInstRow.Add(new InstrumentRow(min, max, currentType, isMerge, instrument, ifold));
                }
            }
            return listInstRow;
        }
    }
}
