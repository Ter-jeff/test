using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RfLib.Dvdc.Reader.DsscSetup
{
    public class DsscSetupSheetReader
    {
        #region  Field
        private const string HeaderDsscSetup = "DsscSetup";
        private const string HeaderPattern = "Pattern";
        private const string HeaderDigSrcEqn = "DigSrcEqn";
        private const string HeaderDigSrcReg = "DigSrcReg";
        private const string HeaderDigSrcAssignment = "DigSrcAssignment";
        private const string HeaderDigSrcPin = "DigSrcPin";
        private const string HeaderDigSrcSampleSize = "DigSrcSampleSize";
        private const string HeaderDigCapPin = "DigCapPin";
        private const string HeaderDigCapSampleSize = "DigCapSampleSize";
        private const string HeaderCusStrDigCapData = "CusStrDigCapData";
        private const string HeaderPatModuleInfo = "PatModuleInfo";
        private readonly Dictionary<string, int> _headerIndexDic;
        #endregion  Field

        #region Constructor
        public DsscSetupSheetReader()
        {
            _headerIndexDic = [];
        }
        #endregion


        public DsscSetupSheet ReadSheet(string fileName)
        {

            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            var dsscsetupSheet = new DsscSetupSheet(sheetName);
            var lines = File.ReadLines(fileName).ToList();
            int maxRowCount = lines.Count;
            bool isbackup = false;
            ReadHeader(lines[0]);
            string setup = "";
            for (int i = 1; i < maxRowCount; i++)
            {
                if (string.IsNullOrEmpty(lines[i]))
                {
                    isbackup = true;
                    continue;
                }
                DsscSetupRow dsscRow = GetDsscRow(lines[i], sheetName, i);

                dsscRow.IsBackup = isbackup;
                if (!string.IsNullOrEmpty(dsscRow.Dsscsetup))
                {
                    setup = dsscRow.Dsscsetup;
                }

                if (!dsscsetupSheet.Setups.TryGetValue(setup, out List<DsscSetupRow>? value))
                {
                    value = [];
                    dsscsetupSheet.Setups.Add(setup, value);
                }

                value.Add(dsscRow);
            }
            return dsscsetupSheet;
        }

        private DsscSetupRow GetDsscRow(string line, string sheetName, int row)
        {
            var dsscRow = new DsscSetupRow
            {
                SheetName = sheetName,
                RowNum = row
            };
            string[] arr = line.Split('\t');
            dsscRow.RowNum = row;
            foreach (KeyValuePair<string, int> item in _headerIndexDic)
            {
                switch (item.Key)
                {
                    case HeaderDsscSetup:
                        dsscRow.Dsscsetup = GetCellText(arr, item.Value);
                        break;
                    case HeaderPattern:
                        dsscRow.Pattern = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigSrcEqn:
                        dsscRow.Digsrceqn = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigSrcReg:
                        dsscRow.Digsrcreg = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigSrcAssignment:
                        dsscRow.Digsrcassignment = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigSrcPin:
                        dsscRow.Digsrcpin = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigSrcSampleSize:
                        dsscRow.Digsrcsamplesize = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigCapPin:
                        dsscRow.Digcappin = GetCellText(arr, item.Value);
                        break;
                    case HeaderDigCapSampleSize:
                        dsscRow.Digcapsamplesize = GetCellText(arr, item.Value);
                        break;
                    case HeaderCusStrDigCapData:
                        dsscRow.Cusstrdigcapdata = GetCellText(arr, item.Value);
                        break;
                    case HeaderPatModuleInfo:
                        dsscRow.PatModule = GetCellText(arr, item.Value);
                        break;
                    default:
                        //dsscRow.Comment = GetCellText(arr,item.Value);
                        break;
                }
            }

            return dsscRow;
        }

        private void ReadHeader(string line)
        {
            string[] arr = line.Split('\t');
            for (int i = 0; i < arr.Length; i++)
            {
                string cellContent = GetCellText(arr, i);
                if (cellContent.Length != 0)
                {
                    _headerIndexDic.Add(cellContent, i);
                }
            }

        }

        public static string GetCellText(string[] line, int column)
        {
            if (column < line.Length)
            {
                return line.ElementAt(column);
            }

            return "";
        }
    }
}
