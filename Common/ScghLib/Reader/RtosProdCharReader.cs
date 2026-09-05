using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace ScghLib.Reader
{
    public partial class RtosProdCharReader(IReadOnlyList<Pin> pinCollect) : MySheetReader<RtosProdCharSheet>
    {
        public const string HeaderPreSetup = "Pre_Setup";
        public const string HeaderTestName = "TestName";
        public const string HeaderType = "Type";
        public const string HeaderUsage = "USAGE";
        public const string HeaderComments = "Comments";
        public const string HeaderCharName = "Char Name";
        public const string HeaderOtherSupplies = "Other Supplies";
        public const string HeaderShmooPin = "ShmooPin";
        public const string HeaderShmooRange = "ShmooRange";
        public const string HeaderCswMeasPin = "CSW_MeasPin";
        public const string HeaderStart = "Start";
        public const string HeaderStop = "Stop";
        public const string HeaderStep = "Step";
        public const string HeaderMethod = "Method";
        public const string HeaderFrequency = "Frequency";

        [GeneratedRegex(HeaderStart, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(HeaderStep, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(HeaderStop, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(HeaderUsage, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex(HeaderCswMeasPin, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();
        [GeneratedRegex(HeaderComments, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex5();

        private readonly Dictionary<string, int> _headerIndexDic = [];
        private readonly Dictionary<int, string> _shmooIndexDic = [];
        private readonly Dictionary<string, int> _paraColumnIndexDic = [];

        private readonly IReadOnlyList<Pin> _pinList = pinCollect;

        private void ReadHeader()
        {
            string cellContent;
            for (int i = 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= ExcelWorksheet.Dimension.End.Column; j++)
                {
                    cellContent = EpplusExtensions.GetCellValue(ExcelWorksheet, i, j);
                    if (cellContent.EqualsIgnoreCase(HeaderTestName))
                    {
                        StartCol = j;
                        StartRow = i;
                        break;
                    }
                }
                if (StartCol != -1)
                {
                    break;
                }
            }

            if (StartCol == -1 || StartRow == -1)
            {
                throw new Exception("Can not find Header: " + HeaderTestName);
            }

            for (int j = StartCol; j <= ExcelWorksheet.Dimension.End.Column; j++)
            {
                cellContent = EpplusExtensions.GetCellValue(ExcelWorksheet, StartRow, j);
                if (cellContent.EqualsIgnoreCase(HeaderOtherSupplies) ||
                    cellContent.EqualsIgnoreCase(HeaderCharName) ||
                    cellContent.EqualsIgnoreCase(HeaderMethod) ||
                    cellContent.EqualsIgnoreCase(HeaderTestName) ||
                    cellContent.EqualsIgnoreCase(HeaderType) ||
                    cellContent.EqualsIgnoreCase(HeaderFrequency) ||
                    cellContent.EqualsIgnoreCase(HeaderCswMeasPin) ||
                    cellContent.EqualsIgnoreCase(HeaderShmooPin) ||
                    cellContent.EqualsIgnoreCase(HeaderShmooRange))
                {
                    _headerIndexDic.Add(cellContent.ToUpper(), j);
                }
                else if (MyRegex().IsMatch(cellContent) ||
                    MyRegex1().IsMatch(cellContent) ||
                    MyRegex2().IsMatch(cellContent))
                {
                    _shmooIndexDic.Add(j, cellContent);
                }
                else if (MyRegex3().IsMatch(cellContent))
                {
                    _headerIndexDic.Add(HeaderUsage.ToUpper(), j);
                }
                else if (MyRegex4().IsMatch(cellContent))
                {
                    _headerIndexDic.Add(HeaderCswMeasPin.ToUpper(), j);
                    _paraColumnIndexDic.Add(cellContent.Replace(" ", "").ToUpper(), j);
                }
                else if (MyRegex5().IsMatch(cellContent))
                {
                    _headerIndexDic.Add(HeaderComments.ToUpper(), j);
                }
                else
                {
                    if (string.IsNullOrEmpty(cellContent.Replace(" ", "")))
                    {
                        continue;
                    }
                    _paraColumnIndexDic.Add(cellContent.Replace(" ", "").ToUpper(), j);
                }
            }
        }

        private RtosProdCharSheet ReadContent()
        {
            var prodSheet = new RtosProdCharSheet();
            int testNameColumn = GetHeaderIndex(HeaderTestName);
            int typeColumn = GetHeaderIndex(HeaderType);
            int usageColumn = GetHeaderIndex(HeaderUsage, false);
            int commentsColumn = GetHeaderIndex(HeaderComments, false);
            int charNameColumn = GetHeaderIndex(HeaderCharName, false);
            int cswColumn = GetHeaderIndex(HeaderCswMeasPin, false);
            if (charNameColumn != -1 || cswColumn != -1)
            {
                prodSheet.IsWithChar = true;
            }

            int otherSupplierColumn = GetHeaderIndex(HeaderOtherSupplies, false);
            int methodColumn = GetHeaderIndex(HeaderMethod, false);
            int shmooPinColumn = GetHeaderIndex(HeaderShmooPin, false);
            int shmooRangeColumn = GetHeaderIndex(HeaderShmooRange, false);

            int frequencyColumn = GetHeaderIndex(HeaderFrequency, false);
            ReadPreSetup(prodSheet);
            for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                var newRow = new RtosProdCharRow
                {
                    TestName = EpplusExtensions.GetCellValue(ExcelWorksheet, i, testNameColumn),
                    TestType = EpplusExtensions.GetCellValue(ExcelWorksheet, i, typeColumn)
                };
                if (usageColumn != -1)
                {
                    newRow.Usage = EpplusExtensions.GetCellValue(ExcelWorksheet, i, usageColumn);
                }

                if (commentsColumn != -1)
                {
                    newRow.Comments = EpplusExtensions.GetCellValue(ExcelWorksheet, i, commentsColumn);
                }

                if (charNameColumn != -1)
                {
                    newRow.CharName = EpplusExtensions.GetCellValue(ExcelWorksheet, i, charNameColumn);
                    newRow.OtherSupplier = EpplusExtensions.GetCellValue(ExcelWorksheet, i, otherSupplierColumn);
                    newRow.Method = EpplusExtensions.GetCellValue(ExcelWorksheet, i, methodColumn);

                    string shmooPinNum = EpplusExtensions.GetCellValue(ExcelWorksheet, i, shmooPinColumn);
                    string shmooRangeNum = EpplusExtensions.GetCellValue(ExcelWorksheet, i, shmooRangeColumn);

                    if (shmooRangeNum.Split(':').Length == 3)
                    {
                        var shmooInfo = new ShmooInfo
                        {
                            ShmooPin = shmooPinNum.Replace(" ", ""),
                            Start = shmooRangeNum.Split(':')[0],
                            Stop = shmooRangeNum.Split(':')[1],
                            Step = shmooRangeNum.Split(':')[2]
                        };
                        newRow.ShmooPinRange.Add(shmooInfo);
                    }
                }
                if (cswColumn != -1)
                {
                    var csw = new CswItem();
                    newRow.Csw = csw;
                    if (frequencyColumn != -1)
                    {
                        csw.Frequency = EpplusExtensions.GetCellValue(ExcelWorksheet, i, frequencyColumn);
                    }

                    csw.MeasPin = EpplusExtensions.GetCellValue(ExcelWorksheet, i, cswColumn);
                    csw.SetShmoo(ExcelWorksheet, i, _shmooIndexDic, StartRow);

                    ReMappingPin(csw, _pinList);
                }

                foreach (KeyValuePair<string, int> headerColumn in _paraColumnIndexDic)
                {
                    string cellContent = EpplusExtensions.GetCellValue(ExcelWorksheet, i, headerColumn.Value);
                    newRow.ParaList.Add(headerColumn.Key.ToUpper(), cellContent);
                }
                if (!string.IsNullOrEmpty(newRow.TestName))
                {
                    prodSheet.RtosRowList.Add(newRow);
                }
            }
            prodSheet.SheetName = ExcelWorksheet.Name;
            return prodSheet;
        }

        public override RtosProdCharSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;
            ReadHeader();
            RtosProdCharSheet outputSheet = ReadContent();
            return outputSheet;
        }

        private void ReadPreSetup(RtosProdCharSheet rtosProdCharSheet)
        {
            for (int i = 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= ExcelWorksheet.Dimension.End.Column; j++)
                {
                    string cellContent = ExcelWorksheet.GetCellValue(i, j);
                    if (cellContent.EqualsIgnoreCase(HeaderPreSetup))
                    {
                        cellContent = ExcelWorksheet.GetCellValue(i + 1, j);
                        rtosProdCharSheet.PreSetup = cellContent;

                        cellContent = ExcelWorksheet.GetCellValue(i + 1, j + 1);
                        rtosProdCharSheet.PreSetupCmd = cellContent;
                    }
                }
            }
        }

        private int GetHeaderIndex(string headerName, bool isMustHeader = true)
        {
            if (_headerIndexDic.ContainsKey(headerName.ToUpper()))
            {
                return _headerIndexDic[headerName.ToUpper()];
            }

            if (isMustHeader)
            {
                throw new Exception("Can not find Column: " + headerName);
            }

            return -1;
        }

        private static void ReMappingPin(CswItem cswItem, IReadOnlyList<Pin> pinCollect)
        {
            foreach (ShmooInfo shmoo in cswItem.Shmoos)
            {
                Pin? target = pinCollect.FirstOrDefault(p => p.PinName.Replace("_", "").EqualsIgnoreCase(shmoo.ShmooPin));
                if (target != null)
                {
                    shmoo.ShmooPin = target.PinName;
                }
            }
        }

    }

    public class RtosProdCharSheet : MySheet
    {
        private const string Function = "Function";
        private const string Char = "Char";

        public string PreSetup { get; set; } = "";
        public string PreSetupCmd { get; set; } = "";
        public List<RtosProdCharRow> RtosRowList { set; get; } = [];

        public bool IsWithChar { set; get; }

        public List<RtosProdCharRow> GetProdRowList()
        {
            return [.. RtosRowList.Where(p => p.TestType.EqualsIgnoreCase(Function))];
        }

        public List<RtosProdCharRow> GetCharRowList()
        {
            return [.. RtosRowList.Where(p => p.TestType.EqualsIgnoreCase(Char))];
        }

        public List<RtosProdCharRow> GetCswList()
        {
            return [.. RtosRowList.Where(p => p.Csw != null)];
        }
    }

    public partial class RtosProdCharRow : MyRow
    {
        public string TestName { get; set; } = string.Empty;
        public string TestType { get; set; } = string.Empty;
        public Dictionary<string, string> ParaList { get; set; } = [];
        public string Usage { get; set; } = string.Empty;
        public string CharName { get; set; } = string.Empty;
        public string OtherSupplier { get; set; } = string.Empty;
        public List<ShmooInfo> ShmooPinRange { get; set; } = [];
        public CswItem? Csw { get; set; } = null;
        public string Method { get; set; } = string.Empty;
        public string Comments { get; set; } = string.Empty;

        public string GetParaValue(string paraName)
        {
            string value = "";
            if (ParaList.ContainsKey(paraName.ToUpper()))
            {
                value = ParaList[paraName.ToUpper()];
            }

            return value;
        }

        public string GetPmgrMode()
        {
            var cmdList = ParaList.Select(p => p.Value).ToList();
            Regex regexRtos = RtosProdCharRowHelpers.MyRegex();
            foreach (string cmd in cmdList)
            {
                Match match = regexRtos.Match(cmd);
                if (match.Success)
                {
                    return match.Groups["PmgrMode"].ToString();
                }
            }
            return "";
        }

        public string GetScenario()
        {
            string cmd = GetParaValue("Cmd4");
            Regex regexRtos = RtosProdCharRowHelpers.MyRegex1();
            Match match = regexRtos.Match(cmd);
            return match.Groups["Scenario"].ToString().Trim();
        }
    }

    public class ShmooInfo
    {
        public bool IsShmoo
        {
            get
            {
                if (string.IsNullOrEmpty(Start) || string.IsNullOrEmpty(Stop))
                {
                    return false;
                }

                return !Start.EqualsIgnoreCase(Stop);
            }
        }
        public string ShmooPin = "";
        public string Start = "";
        public string Stop = "";
        public string Step = "";
    }

    public class CswItem
    {
        public string Frequency = string.Empty;
        public string MeasPin = string.Empty;
        public List<ShmooInfo> Shmoos = [];

        public string GetCswName()
        {
            //RTOS_CSWChar_<Freq>_<CSW Meas Pin>
            var nameList = new List<string> { "RTOS", "CSWChar" };
            if (!string.IsNullOrEmpty(Frequency))
            {
                nameList.Add(Frequency.Replace(".", "p") + "MHz");
            }

            if (!string.IsNullOrEmpty(MeasPin))
            {
                nameList.Add(MeasPin.Replace("_", ""));
            }

            return string.Join("_", nameList);
        }

        public void SetShmoo(ExcelWorksheet excelWorksheet, int rowIndex, Dictionary<int, string> shmooDictionary, int headerRow)
        {
            foreach (KeyValuePair<int, string> index in shmooDictionary)
            {
                if (index.Value.EqualsIgnoreCase("start"))
                {
                    var shmoo = new ShmooInfo
                    {
                        ShmooPin = EpplusExtensions.GetCellValue(excelWorksheet, headerRow - 1, index.Key),
                        Start = EpplusExtensions.GetCellValue(excelWorksheet, rowIndex, index.Key),
                        Stop = EpplusExtensions.GetCellValue(excelWorksheet, rowIndex, index.Key + 1),
                        Step = EpplusExtensions.GetCellValue(excelWorksheet, rowIndex, index.Key + 2)
                    };
                    Shmoos.Add(shmoo);
                }
            }
        }
    }
}
