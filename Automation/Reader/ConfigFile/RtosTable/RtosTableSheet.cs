using System;
using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.Reader.ConfigFile.RtosTable
{
    public class RtosTableSheet
    {
        #region Field

        private const string Sheetname = "SheetName";
        private const string Funcname = "FuncName";
        private const string Argname = "ArgName";
        private const string Arginfo = "ArgInfo";

        private const string Pinname = "PinName";
        private const string Pinvalue = "PinValue";

        private const string Binaryfilename = "BinaryFileName";

        private const string Rtoslevel = "RtosLevel";
        private const string Parameter = "Parameter";
        private const string Value = "Value";

        private const string Measmode = "MeasMode";
        private const string Measpattern = "MeasPattern";

        #endregion

        #region Properity

        public List<RtosTableArgRow> ArgRows { get; } = new List<RtosTableArgRow>();
        public Dictionary<string, string> PinRows { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string BinFileName { get; private set; } = "";
        public List<RtosLevelRow> LevelRows { get; } = new List<RtosLevelRow>();
        public Dictionary<string, string> MeasRows { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        #endregion

        public static RtosTableSheet LoadConfig(ExcelWorksheet worksheet)
        {
            var rtostable = new RtosTableSheet();
            try
            {

                int sheetnameRow = 0, sheetnameCol = 0, funcnameCol = 0, argnameCol = 0, arginfoCol = 0;
                int pinnameRow = 0, pinnameCol = 0, pinvalCol = 0;
                int binaryfilenameRow = 0, binaryfilenameCol = 0;
                int levelRow = 0, levelCol = 0, parameterCol = 0, valueCol = 0;
                int measmodeRow = 0, measmodeCol = 0, measpatternCol = 0;

                ReadHeaders(worksheet,
                    ref sheetnameRow, ref sheetnameCol, ref funcnameCol, ref argnameCol, ref arginfoCol,
                    ref pinnameRow, ref pinnameCol, ref pinvalCol,
                    ref binaryfilenameRow, ref binaryfilenameCol,
                    ref levelRow, ref levelCol, ref parameterCol, ref valueCol,
                    ref measmodeRow, ref measmodeCol, ref measpatternCol);

                ReadArgumentInfo(worksheet, rtostable, sheetnameRow, sheetnameCol, funcnameCol, argnameCol, arginfoCol);
                ReadPinInfo(worksheet, rtostable, pinnameRow, pinnameCol, pinvalCol);
                ReadBinaryInfo(worksheet, rtostable, binaryfilenameRow, binaryfilenameCol);
                ReadLevelInfo(worksheet, rtostable, levelRow, levelCol, parameterCol, valueCol);
                ReadMeasInfo(worksheet, rtostable, measmodeRow, measmodeCol, measpatternCol);
            }
            catch (Exception e)
            {
                throw new Exception("Error occurs during load Rtos Table Config: " + e.StackTrace);
            }

            return rtostable;
        }

        private static void ReadHeaders(ExcelWorksheet worksheet,
            ref int sheetnameRow, ref int sheetnameCol, ref int funcnameCol, ref int argnameCol, ref int arginfoCol,
            ref int pinnameRow, ref int pinnameCol, ref int pinvalCol,
            ref int binaryfilenameRow, ref int binaryfilenameCol,
            ref int levelRow, ref int levelCol, ref int parameterCol, ref int valueCol,
            ref int measmodeRow, ref int measmodeCol, ref int measpatternCol)
        {
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    string cellContent = EpplusExtensions.GetCellValue(worksheet, i, j);
                    if (MatchArgHeader(cellContent, i, j,
                        ref sheetnameRow, ref sheetnameCol, ref funcnameCol, ref argnameCol, ref arginfoCol))
                    {
                        continue;
                    }
                    if (MatchPinAndBinaryHeader(cellContent, i, j,
                        ref pinnameRow, ref pinnameCol, ref pinvalCol, ref binaryfilenameRow, ref binaryfilenameCol))
                    {
                        continue;
                    }
                    MatchLevelAndMeasHeader(cellContent, i, j,
                        ref levelRow, ref levelCol, ref parameterCol, ref valueCol,
                        ref measmodeRow, ref measmodeCol, ref measpatternCol);
                }
            }
        }

        private static bool MatchArgHeader(string cellContent, int i, int j,
            ref int sheetnameRow, ref int sheetnameCol, ref int funcnameCol, ref int argnameCol, ref int arginfoCol)
        {
            if (cellContent.Equals(Sheetname, StringComparison.OrdinalIgnoreCase))
            {
                sheetnameRow = i;
                sheetnameCol = j;
                return true;
            }
            if (cellContent.Equals(Funcname, StringComparison.OrdinalIgnoreCase))
            {
                funcnameCol = j;
                return true;
            }
            if (cellContent.Equals(Argname, StringComparison.OrdinalIgnoreCase))
            {
                argnameCol = j;
                return true;
            }
            if (cellContent.Equals(Arginfo, StringComparison.OrdinalIgnoreCase))
            {
                arginfoCol = j;
                return true;
            }
            return false;
        }

        private static bool MatchPinAndBinaryHeader(string cellContent, int i, int j,
            ref int pinnameRow, ref int pinnameCol, ref int pinvalCol, ref int binaryfilenameRow, ref int binaryfilenameCol)
        {
            if (cellContent.Equals(Pinname, StringComparison.OrdinalIgnoreCase))
            {
                pinnameRow = i;
                pinnameCol = j;
                return true;
            }
            if (cellContent.Equals(Pinvalue, StringComparison.OrdinalIgnoreCase))
            {
                pinvalCol = j;
                return true;
            }
            if (cellContent.Equals(Binaryfilename, StringComparison.OrdinalIgnoreCase))
            {
                binaryfilenameRow = i;
                binaryfilenameCol = j;
                return true;
            }
            return false;
        }

        private static void MatchLevelAndMeasHeader(string cellContent, int i, int j,
            ref int levelRow, ref int levelCol, ref int parameterCol, ref int valueCol,
            ref int measmodeRow, ref int measmodeCol, ref int measpatternCol)
        {
            if (cellContent.Equals(Rtoslevel, StringComparison.OrdinalIgnoreCase))
            {
                levelRow = i;
                levelCol = j;
            }
            else if (cellContent.Equals(Parameter, StringComparison.OrdinalIgnoreCase))
            {
                parameterCol = j;
            }
            else if (cellContent.Equals(Value, StringComparison.OrdinalIgnoreCase))
            {
                valueCol = j;
            }
            else if (cellContent.Equals(Measmode, StringComparison.OrdinalIgnoreCase))
            {
                measmodeRow = i;
                measmodeCol = j;
            }
            else if (cellContent.Equals(Measpattern, StringComparison.OrdinalIgnoreCase))
            {
                measpatternCol = j;
            }
        }

        private static void ReadArgumentInfo(ExcelWorksheet worksheet, RtosTableSheet rtostable,
            int sheetnameRow, int sheetnameCol, int funcnameCol, int argnameCol, int arginfoCol)
        {
            if (sheetnameRow <= 0)
            {
                return;
            }

            for (int i = sheetnameRow + 1; ; i++)
            {
                if (EpplusExtensions.GetCellValue(worksheet, i, 1) == "")
                {
                    break;
                }

                rtostable.ArgRows.Add(new RtosTableArgRow(
                    EpplusExtensions.GetCellValue(worksheet, i, sheetnameCol),
                    EpplusExtensions.GetCellValue(worksheet, i, funcnameCol),
                    EpplusExtensions.GetCellValue(worksheet, i, argnameCol),
                    EpplusExtensions.GetCellValue(worksheet, i, arginfoCol)
                    ));
            }
        }

        private static void ReadPinInfo(ExcelWorksheet worksheet, RtosTableSheet rtostable,
            int pinnameRow, int pinnameCol, int pinvalCol)
        {
            if (pinnameRow <= 0)
            {
                return;
            }

            for (int i = pinnameRow + 1; ; i++)
            {
                if (EpplusExtensions.GetCellValue(worksheet, i, 1) == "")
                {
                    break;
                }

                rtostable.PinRows.Add(
                    EpplusExtensions.GetCellValue(worksheet, i, pinnameCol),
                    EpplusExtensions.GetCellValue(worksheet, i, pinvalCol));
            }
        }

        private static void ReadBinaryInfo(ExcelWorksheet worksheet, RtosTableSheet rtostable,
            int binaryfilenameRow, int binaryfilenameCol)
        {
            if (binaryfilenameRow <= 0)
            {
                return;
            }

            for (int i = binaryfilenameRow + 1; ; i++)
            {
                if (EpplusExtensions.GetCellValue(worksheet, i, 1) == "")
                {
                    break;
                }

                rtostable.BinFileName = EpplusExtensions.GetCellValue(worksheet, i, binaryfilenameCol);
            }
        }

        private static void ReadLevelInfo(ExcelWorksheet worksheet, RtosTableSheet rtostable,
            int levelRow, int levelCol, int parameterCol, int valueCol)
        {
            if (levelRow <= 0)
            {
                return;
            }

            for (int i = levelRow + 1; ; i++)
            {
                if (EpplusExtensions.GetCellValue(worksheet, i, 1) == "")
                {
                    break;
                }

                string pinName = EpplusExtensions.GetCellValue(worksheet, i, levelCol);
                string parameter = EpplusExtensions.GetCellValue(worksheet, i, parameterCol);
                string value = EpplusExtensions.GetCellValue(worksheet, i, valueCol);
                RtosLevelRow targetRow = rtostable.LevelRows.Find(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                if (targetRow == null)
                {
                    rtostable.LevelRows.Add(new RtosLevelRow { PinName = pinName, Parameters = new Dictionary<string, string>() });
                    targetRow = rtostable.LevelRows.Find(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                }
                if (!string.IsNullOrEmpty(parameter) && !string.IsNullOrEmpty(value))
                {
                    targetRow.Parameters.Add(parameter, value);
                }
            }
        }

        private static void ReadMeasInfo(ExcelWorksheet worksheet, RtosTableSheet rtostable,
            int measmodeRow, int measmodeCol, int measpatternCol)
        {
            if (measmodeRow <= 0)
            {
                return;
            }

            for (int i = measmodeRow + 1; ; i++)
            {
                if (EpplusExtensions.GetCellValue(worksheet, i, 1) == "")
                {
                    break;
                }

                rtostable.MeasRows.Add(
                    EpplusExtensions.GetCellValue(worksheet, i, measmodeCol),
                    EpplusExtensions.GetCellValue(worksheet, i, measpatternCol));
            }
        }

    }
    public class RtosLevelRow
    {
        public string PinName { get; set; }

        public Dictionary<string, string> Parameters { get; set; }
    }
    public class RtosTableArgRow
    {
        #region Properity

        public string SheetName { get; }
        public string FuncName { get; }
        public string ArgName { get; }
        public string ArgInfo { get; }

        #endregion

        #region Constructor

        public RtosTableArgRow(string sheetname, string funcname, string argname, string arginfo)
        {
            SheetName = sheetname;
            FuncName = funcname;
            ArgName = argname;
            ArgInfo = arginfo;
        }

        #endregion

    }

}
