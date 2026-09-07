using System;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;

namespace TestPlanLib.HardIpDc.Business
{
    public class HardIpDcReader
    {
        #region Filed

        private const int MaxSearchColumn = 10;
        private const int MaxSearcRow = 10;
        public const string HeaderNv = "NV";
        public const string HeaderNvValt = "NV(Valt)";
        public const string HeaderHvRatio = "HV_Ratio";
        public const string HeaderLvRatio = "LV_Ratio";
        public const string HeaderIfold = "Ifold";
        public const string HeaderVil = "Vil";
        public const string HeaderVih = "Vih";
        public const string HeaderVol = "Vol";
        public const string HeaderVoh = "Voh";
        public const string HeaderIol = "Iol";
        public const string HeaderIoh = "Ioh";
        public const string HeaderVt = "Vt";
        public const string HeaderVcl = "Vcl";
        public const string HeaderVch = "Vch";
        public const string HeaderDriveMode = "DriverMode";
        public const string HeaderVicm = "Vicm";
        public const string HeaderVid = "Vid";
        public const string HeaderVod = "Vod";

        private ExcelWorksheet? _inputWorksheet;
        private readonly HardIpDcSheet _outPutSheet = new();
        private int _startColumn;
        private int _startRow;
        private readonly Headers _headers = new();
        #endregion

        #region Read Flow
        public HardIpDcSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            try
            {
                if (excelWorksheet == null)
                {
                    return null;
                }

                _inputWorksheet = excelWorksheet;

                ReadHeader();

                ReadData();
            }
            catch (Exception)
            {
                return null;
            }

            return _outPutSheet;
        }
        #endregion

        #region Member function

        private void ReadHeader()
        {
            string cellContent;
            bool hasFind = false;
            //locate the first header
            for (int i = 1; i <= MaxSearcRow; i++)
            {
                for (int j = 1; j <= MaxSearchColumn; j++)
                {
                    cellContent = EpplusExtensions.GetCellValue(_inputWorksheet!, i, j);
                    if (cellContent.EqualsIgnoreCase(HeaderNv))
                    {
                        _startRow = i;
                        _startColumn = j;
                        hasFind = true;
                        break;
                    }
                }
                if (hasFind)
                {
                    break;
                }
            }

            //Get All header columns
            for (int i = _startColumn; i <= _inputWorksheet!.Dimension.End.Column; i++)
            {
                cellContent = EpplusExtensions.GetCellValue(_inputWorksheet!, _startRow, i);
                if (cellContent.Length == 0)
                {
                    break;
                }
                _headers.AddHeaderItem(cellContent, i);
            }
        }

        private void ReadData()
        {
            int nvColumn = _headers.GetHeaderIndex(HeaderNv);
            int nvValtColumn = _headers.GetHeaderIndex(HeaderNvValt, false);
            int hvRatioColumn = _headers.GetHeaderIndex(HeaderHvRatio);
            int lvRatioColumn = _headers.GetHeaderIndex(HeaderLvRatio);
            int ifoldColumn = _headers.GetHeaderIndex(HeaderIfold);
            int vilColumn = _headers.GetHeaderIndex(HeaderVil);
            int vihColumn = _headers.GetHeaderIndex(HeaderVih);
            int volColumn = _headers.GetHeaderIndex(HeaderVol);
            int vohColumn = _headers.GetHeaderIndex(HeaderVoh);
            int iolColumn = _headers.GetHeaderIndex(HeaderIol);
            int iohColumn = _headers.GetHeaderIndex(HeaderIoh);
            int vtColumn = _headers.GetHeaderIndex(HeaderVt);
            int vclColumn = _headers.GetHeaderIndex(HeaderVcl);
            int vchColumn = _headers.GetHeaderIndex(HeaderVch);
            int driverModeColumn = _headers.GetHeaderIndex(HeaderDriveMode);
            int vicmColum = _headers.GetHeaderIndex(HeaderVicm);
            int vidColumn = _headers.GetHeaderIndex(HeaderVid);
            int vodColumn = _headers.GetHeaderIndex(HeaderVod);

            string categoryName = EpplusExtensions.GetCellValue(_inputWorksheet!, _startRow, nvColumn - 1);
            HardIpCategoryDef categoryDef = new HardIpCategoryDef(categoryName);
            for (int i = _startRow + 1; i <= _inputWorksheet!.Dimension.End.Row; i++)
            {
                string cellContent = EpplusExtensions.GetCellValue(_inputWorksheet!, i, nvColumn);
                if (cellContent.EqualsIgnoreCase(HeaderNv))
                {
                    categoryDef.DcCategory = GetDcCategoryName(categoryName, categoryDef);
                    categoryDef.LevelSheet = categoryDef.GetLevelName();
                    _outPutSheet.Rows.Add(categoryDef);
                    //Start a new Category
                    categoryName = EpplusExtensions.GetCellValue(_inputWorksheet!, i, nvColumn - 1);
                    categoryDef = new HardIpCategoryDef(categoryName);
                }
                else
                {
                    //add pin definitions
                    HardIpDcRow row = new HardIpDcRow
                    {
                        PinName = EpplusExtensions.GetCellValue(_inputWorksheet!, i, nvColumn - 1),
                        Nv = EpplusExtensions.GetCellValue(_inputWorksheet!, i, nvColumn)
                    };
                    if (nvValtColumn != -1)
                    {
                        row.NvValt = EpplusExtensions.GetCellValue(_inputWorksheet!, i, nvValtColumn);
                    }

                    row.HvRatio = EpplusExtensions.GetCellValue(_inputWorksheet!, i, hvRatioColumn);
                    row.LvRatio = EpplusExtensions.GetCellValue(_inputWorksheet!, i, lvRatioColumn);
                    row.Ifold = EpplusExtensions.GetCellValue(_inputWorksheet!, i, ifoldColumn);
                    row.Vil = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vilColumn);
                    row.Vih = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vihColumn);
                    row.Vol = EpplusExtensions.GetCellValue(_inputWorksheet!, i, volColumn);
                    row.Voh = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vohColumn);
                    row.Iol = EpplusExtensions.GetCellValue(_inputWorksheet!, i, iolColumn);
                    row.Ioh = EpplusExtensions.GetCellValue(_inputWorksheet!, i, iohColumn);
                    row.Vt = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vtColumn);
                    row.Vcl = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vclColumn);
                    row.Vch = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vchColumn);
                    row.DriverMode = EpplusExtensions.GetCellValue(_inputWorksheet!, i, driverModeColumn);
                    row.Vicm = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vicmColum);
                    row.Vid = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vidColumn);
                    row.Vod = EpplusExtensions.GetCellValue(_inputWorksheet!, i, vodColumn);
                    row.RowNum = i.ToString();
                    if (row.PinName?.Length != 0)
                    {
                        categoryDef.DataRows.Add(row);
                    }
                }
            }
            categoryDef.DcCategory = GetDcCategoryName(categoryName, categoryDef);
            categoryDef.LevelSheet = categoryDef.GetLevelName();
            if (!string.IsNullOrEmpty(categoryName))
            {
                _outPutSheet.Rows.Add(categoryDef);
            }
        }

        private static string GetDcCategoryName(string categoryName, HardIpCategoryDef hardIpCategoryDef)
        {
            //Change by Kimi at 2020/07/10 , if all Nv Hv Lv value are empty, use default DC Category
            if (categoryName.EqualsIgnoreCase("Scan") ||
                categoryName.EqualsIgnoreCase("Mbist"))
            {
                return categoryName;
            }
            return hardIpCategoryDef.DataRows.Any(p => p.Nv?.Length != 0 && p.HvRatio?.Length != 0 && p.LvRatio?.Length != 0)
                ? "HardIP_" + categoryName
                : "";

            //return categoryName.Equals("Scan", StringComparison.OrdinalIgnoreCase) ||
            //    categoryName.Equals("Mbist", StringComparison.OrdinalIgnoreCase)
            //    ? categoryName : "HardIP_" + categoryName;
        }

        #endregion
    }
}
