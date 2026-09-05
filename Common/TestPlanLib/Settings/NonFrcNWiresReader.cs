using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.Settings
{
    public partial class NonFrcNWiresReader
    {
        #region Field
        private const string PortName = "Port Name";
        private const string ProtocalType = "Protocol Type";
        private const string FunctionName = "Function Name";
        private const string DeviecPinName = "Device Pin Name";

        [GeneratedRegex(ProtocalType, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(FunctionName, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(DeviecPinName, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();

        private ExcelWorksheet? _worksheet;
        private int _startRow = 1;
        private int _startColumn = 1;

        private int _portName = -1;
        private int _protocalType = -1;
        private int _functionName = -1;
        private int _deviecPinName = -1;
        private List<NonFrcNWires> _nonFrcNWiresList = [];
        #endregion

        #region Member Function

        public List<NonFrcNWires> ReadFlow(ExcelWorksheet excelWorksheet)
        {
            try
            {
                _worksheet = excelWorksheet;

                ReadHeader();

                ReadData();

            }
            catch (Exception e)
            {

                throw new Exception("Error occurs during Reading nWire setting file: " + e.StackTrace);

            }
            return _nonFrcNWiresList;
        }

        private void ReadHeader()
        {
            string header;
            bool hasFind = false;
            for (int i = 1; i <= _worksheet!.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= _worksheet!.Dimension.End.Column; j++)
                {
                    header = EpplusExtensions.GetCellValue(_worksheet!, i, j);
                    if (header.EqualsIgnoreCase(PortName))
                    {
                        hasFind = true;
                        _startRow = i;
                        _startColumn = j;
                        break;
                    }
                }
                if (hasFind)
                {
                    break;
                }
            }
            _portName = _startRow;

            for (int i = _startColumn; i <= _worksheet!.Dimension.End.Column; i++)
            {
                header = EpplusExtensions.GetCellValue(_worksheet!, _startRow, i);
                if (header.Length == 0)
                {
                    break;
                }

                if (MyRegex().IsMatch(header))
                {
                    _protocalType = i;
                }
                else if (MyRegex1().IsMatch(header))
                {
                    _functionName = i;
                }
                else if (MyRegex2().IsMatch(header))
                {
                    _deviecPinName = i;
                }
            }
        }

        private void ReadData()
        {
            _nonFrcNWiresList = [];
            for (int i = _startRow + 1; i <= _worksheet!.Dimension.End.Row; i++)
            {
                NonFrcNWires row = new NonFrcNWires();
                if (_portName != -1)
                {
                    row.PortName = EpplusExtensions.GetCellValue(_worksheet!, i, _portName);
                }

                if (_protocalType != -1)
                {
                    row.ProtocalType = EpplusExtensions.GetCellValue(_worksheet!, i, _protocalType);
                }

                if (_functionName != -1)
                {
                    row.FunctionName = EpplusExtensions.GetCellValue(_worksheet!, i, _functionName);
                }

                if (_deviecPinName != -1)
                {
                    row.DeviecPinName = EpplusExtensions.GetCellValue(_worksheet!, i, _deviecPinName);
                }

                _nonFrcNWiresList.Add(row);
            }
        }
        #endregion
    }

    public class NonFrcNWires
    {
        #region Field
        public const string Consetting = "setting";
        #endregion

        #region Properity
        public string PortName { set; get; } = "";
        public string ProtocalType { set; get; } = "";
        public string FunctionName { set; get; } = "";
        public string DeviecPinName { set; get; } = "";
        #endregion

        public static string CreateSettingName(XmlNode xmlNode, string type)
        {
            if (xmlNode != null && xmlNode.Attributes != null)
            {
                foreach (XmlNode childNodes in xmlNode.ChildNodes)
                {
                    if (childNodes.Name.EqualsIgnoreCase(Consetting))
                    {
                        if (childNodes.Attributes != null)
                        {
                            return childNodes.Attributes[type]!.Value;
                        }
                    }
                }
            }

            return "";
        }

        public static string CreateSettingdefaultValue(XmlNode xmlNode)
        {
            if (xmlNode != null && xmlNode.Attributes != null)
            {
                foreach (XmlNode childNodes in xmlNode.ChildNodes)
                {
                    if (childNodes.Name.EqualsIgnoreCase(Consetting))
                    {
                        if (childNodes.Attributes != null)
                        {
                            return childNodes.Attributes["defaultValue"]!.Value;
                        }
                    }
                }
            }

            return "";
        }
    }
}
