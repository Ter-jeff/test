using System.Linq;
using System.Text.RegularExpressions;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    internal partial class EfuseReadLine : XLine
    {
        public static readonly Regex RegexBinary = BinaryRegex();
        public static readonly Regex RegexHex = HexRegex();

        [GeneratedRegex("b[01]+", RegexOptions.IgnoreCase)]
        private static partial Regex BinaryRegex();

        [GeneratedRegex("x[0-9A-Fa-f]+", RegexOptions.IgnoreCase)]
        private static partial Regex HexRegex();

        //Site(0) EFUSE Read Values Bank_ecid(MSB)00000:00035(LSB) CP1 lot_id = E9S744 = b001110001001011100000111000100000100
        //Site(0) EFUSE Read Values Bank_ecid(MSB)00036:00040(LSB) CP1 wafer_id = x0000000000000017 = 23 = b10111
        //Site(0) EFUSE Read Values Bank_ecid(MSB)00041:00046(LSB) CP1 x_coordinate = x0000000000000010 = 16 = b010000
        //Site(0) EFUSE Read Values Bank_ecid(MSB)00047:00052(LSB) CP1 y_coordinate = x0000000000000011 = 17 = b010001
        //Site(0) EFUSE Read Values Bank_ecid(MSB)00053:00063(LSB) CP1 bank_ecid_Spare0 = x0000000000000000 = 0 = b00000000000
        //Site(0) EFUSE Read Values Bank_ecid(MSB)00255:00224(LSB) CP1 ecid_crc = x0000000069D46EA1 = 1775529633 = b01101001110101000110111010100001
        public EfuseReadRow ConvertEfuseReadRow()
        {
            var row = new EfuseReadRow();
            string[] arr = Line.Split('=');
            row.Site = GetSite3();
            row.Name = arr[0].Trim().Split(' ').Last();
            string value = arr[1].Trim();
            if (RegexBinary.IsMatch(value))
            {
                row.Value = BinaryToDecimal(value[1..]).ToString();
            }
            else if (RegexHex.IsMatch(value))
            {
                row.Value = HexToDecimal(value[1..]).ToString();
            }
            else
            {
                row.Value = value;
            }
            row.Line = this;
            return row;
        }
    }
}
