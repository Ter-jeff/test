using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public partial class EfuseCfgTable
    {
        [GeneratedRegex("Condition", RegexOptions.IgnoreCase)]
        private static partial Regex ConditionRegex();

        [GeneratedRegex("MSB BIT", RegexOptions.IgnoreCase)]
        private static partial Regex MsbBitRegex();

        [GeneratedRegex("LSB BIT", RegexOptions.IgnoreCase)]
        private static partial Regex LsbBitRegex();

        [GeneratedRegex("Bit Width", RegexOptions.IgnoreCase)]
        private static partial Regex BitWidthRegex();

        [GeneratedRegex("programming stage", RegexOptions.IgnoreCase)]
        private static partial Regex ProgrammingStageRegex();

        public int ConditionIdx = -1;
        public int MsbBitIdx = -1;
        public int LsbBitIdx = -1;
        public int BitWidthIdx = -1;
        public int PrgStageIdx = -1;
        public Dictionary<string, int> Scenario = [];
        public List<List<string>> CfgRows = [];
        public string TableName = "";
        public List<string> TitleList = [];
        public void Read(ExcelWorksheet excelWorksheet)
        {
            TableName = excelWorksheet.Name;
            string? line;

            int startrow = 1;
            int endrow = excelWorksheet.Dimension.End.Row;
            while (startrow <= endrow)
            {
                line = excelWorksheet.GetCellLine(startrow);
                startrow++;
                if (ConditionRegex().IsMatch(line))
                {
                    ReadHeader(line);
                    break;
                }

            }

            while (startrow <= endrow)
            {
                line = excelWorksheet.GetCellLine(startrow).ToUpper();
                startrow++;
                List<string> vLine = [.. line.Split('\t')];
                CfgRows.Add(vLine);
            }

            CfgRows = [.. CfgRows.Where(x => x.Count > ConditionIdx).Where(x => !string.IsNullOrEmpty(x.ElementAt(ConditionIdx)))];
        }

        public void ReadTxt(string fileName)
        {
            var reader = new StreamReader(fileName);
            TableName = Path.GetFileNameWithoutExtension(fileName);
            string? line;

            while ((line = reader.ReadLine()) != null)
            {
                if (ConditionRegex().IsMatch(line))
                {
                    ReadHeader(line);
                    break;
                }

            }

            while ((line = reader.ReadLine()) != null)
            {
                line = line.ToUpper();
                List<string> vLine = [.. line.Split('\t')];
                CfgRows.Add(vLine);
            }
            CfgRows = [.. CfgRows.Where(x => x.Count > ConditionIdx).Where(x => !string.IsNullOrEmpty(x.ElementAt(ConditionIdx)))];
            reader.Close();
        }

        /*    EFUSE_CONFIG_MAIN_A																EFUSE_CONFIG_MAIN_B											
Condition	MSB BIT	LSB BIT	Bit Width	programming stage	A00	A01	A02	A03	A04	A05	A06	A07	A09	A12	CommentA	B00	B01	B02	B03	B04	B05	B06	B07	B09	B12	CommentB	end
    */
        private void ReadHeader(string line)
        {
            TitleList.Clear();
            string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
            //lineSpt = ReArrangeLin(lineSpt, lineLen);
            for (int i = 0; i < lineSpt.Length; i++)
            {
                if (ConditionRegex().IsMatch(lineSpt[i]))
                {
                    ConditionIdx = i;
                }
                else if (MsbBitRegex().IsMatch(lineSpt[i]))
                {
                    MsbBitIdx = i;
                }
                else if (LsbBitRegex().IsMatch(lineSpt[i]))
                {
                    LsbBitIdx = i;
                }
                else if (BitWidthRegex().IsMatch(lineSpt[i]))
                {
                    BitWidthIdx = i;
                }
                else if (ProgrammingStageRegex().IsMatch(lineSpt[i]))
                {
                    PrgStageIdx = i;
                }
                else if (!string.IsNullOrEmpty(lineSpt[i].Trim()))
                {
                    Scenario.Add(lineSpt[i], i);
                }
            }
            TitleList = [.. lineSpt];
        }

    }
}
