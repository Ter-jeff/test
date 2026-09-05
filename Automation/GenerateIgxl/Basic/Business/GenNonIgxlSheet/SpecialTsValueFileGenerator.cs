using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenDc;
using Automation.Static;

using TestPlanLib.DataStruct;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenNonIgxlSheet
{
    public class SpecialTsValueFileGenerator
    {
        private readonly List<TestSettingData> _testSettingSheetlst;
        private readonly List<string> _specialCategorylst = new List<string>();
        private readonly Regex _specialCategoryReg = DcSpecGenerator.Reg;
        private readonly string _outputFolder;

        public SpecialTsValueFileGenerator(List<TestSettingData> testSettingSheetlst, string outputFolder)
        {
            _testSettingSheetlst = testSettingSheetlst;
            _outputFolder = outputFolder;
        }

        public List<string> GenerateSpecialTsValueFile()
        {
            var filelst = new List<string>();
            foreach (TestSettingData ts in _testSettingSheetlst)
            {
                TestSettingData editDt = EditTestSettingDt(ts);
                if (editDt != null)
                {
                    string specialTsValueFilePath = Path.Combine(_outputFolder, ts.SheetName + "_BinCut") + ".TXT";
                    WriteToFile(editDt, specialTsValueFilePath);
                    filelst.Add(specialTsValueFilePath);
                }
            }
            return filelst;
        }

#nullable enable
        public TestSettingData? EditTestSettingDt(TestSettingData ts)
#nullable restore
        {
            _specialCategorylst.Clear();
            List<DcCategoryName> categoryNameList = ts.DcCategorys;
            //Find categorys that contains special value
            foreach (TestSettingRow row in ts.DataRows)
            {
                foreach (DcCategoryValue dcValue in row.DcCategoryValues)
                {
                    string catgeoryName = categoryNameList.Find(x => x.ColumnIndex == dcValue.ColumnIndex).CategoryName.ToUpper();
                    if (_specialCategorylst.Contains(catgeoryName))
                    {
                        continue;
                    }

                    string strContent = dcValue.Nv.Value;
                    if (_specialCategoryReg.IsMatch(strContent))
                    {
                        if (!_specialCategorylst.Contains(catgeoryName))
                        {
                            _specialCategorylst.Add(catgeoryName);
                        }

                        continue;
                    }

                    strContent = dcValue.Hv.Value;
                    if (_specialCategoryReg.IsMatch(strContent))
                    {
                        if (!_specialCategorylst.Contains(catgeoryName))
                        {
                            _specialCategorylst.Add(catgeoryName);
                        }

                        continue;
                    }

                    strContent = dcValue.Lv.Value;
                    if (_specialCategoryReg.IsMatch(strContent))
                    {
                        if (!_specialCategorylst.Contains(catgeoryName))
                        {
                            _specialCategorylst.Add(catgeoryName);
                        }
                    }
                }
            }

            if (_specialCategorylst.Count == 0)
            {
                return null;
            }

            return ts;
        }

        private string GetBinCutSheetCount()
        {
            var lstCount = new List<string>();
            if (EpWorkbook.BinCutWorkbook == null)
            {
                return string.Empty;
            }

            if (EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.Binning] != null)
            {
                lstCount.Add("1");
            }

            if (EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BinningBinX] != null)
            {
                lstCount.Add("2");
            }

            if (EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BinningBinY] != null)
            {
                lstCount.Add("3");
            }

            return string.Join(",", lstCount);
        }

        /// <summary>
        /// Write Datatable to TXT File
        /// </summary>
        /// <param name="pTable">Data table</param>
        /// <param name="filePath">FileName</param>
        private void WriteToFile(TestSettingData pTable, string filePath)
        {
            StreamWriter sw = null;
            var re = new Regex(@"(?<pMode>\w+)[ _]+(?<others>.*)");

            if (!Directory.Exists(_outputFolder))
            {
                Directory.CreateDirectory(_outputFolder);
            }

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            var aa = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            sw = new StreamWriter(aa);

            Dictionary<string, List<DcCategoryName>> specialColumns = GetSpecialColumnNumber(pTable);
            //write category name line
            string bcCountStr = GetBinCutSheetCount();
            sw.Write("Rev:\t" + pTable.TestSettingVersion + "\t\t" + "Bin Cut List=\t" + bcCountStr + "\r\n");
            for (int i = 0; i < _specialCategorylst.Count; i++)
            {
                if (i == _specialCategorylst.Count - 1)
                {
                    List<DcCategoryName> dcNames = specialColumns[_specialCategorylst[i]];
                    for (int j = 0; j < dcNames.Count; j++)
                    {
                        if (j == dcNames.Count - 1)
                        {
                            sw.Write(dcNames[j].CategoryName + "\r\n");
                        }
                        else
                        {
                            sw.Write(dcNames[j].CategoryName + "\t");
                        }
                    }
                }
                else
                {
                    if (i == 0)
                    {
                        sw.Write("Category\t");
                    }

                    foreach (DcCategoryName dcName in specialColumns[_specialCategorylst[i]])
                    {
                        sw.Write(dcName.CategoryName + "\t");
                    }
                }
            }

            //write category value type line
            sw.Write("PinName\t");
            for (int i = 0; i < _specialCategorylst.Count; i++)
            {
                if (i == _specialCategorylst.Count - 1)
                {
                    List<DcCategoryName> dcNames = specialColumns[_specialCategorylst[i]];
                    for (int j = 0; j < dcNames.Count; j++)
                    {
                        if (j == dcNames.Count - 1)
                        {
                            sw.Write(dcNames[j].ValueType + "\r\n");
                        }
                        else
                        {
                            sw.Write(dcNames[j].ValueType + "\t");
                        }
                    }
                }
                else
                {
                    foreach (DcCategoryName dcName in specialColumns[_specialCategorylst[i]])
                    {
                        sw.Write(dcName.ValueType + "\t");
                    }
                }
            }

            foreach (TestSettingRow dataRow in pTable.DataRows)
            {
                sw.Write(dataRow.PowerPinName + "\t");
                for (int i = 0; i < _specialCategorylst.Count; i++)
                {
                    List<DcCategoryName> dcNames = specialColumns[_specialCategorylst[i]];
                    int columnIndex = dcNames[0].ColumnIndex;
                    for (int j = 0; j < dcNames.Count; j++)
                    {
                        string value;
                        DcCategoryValue dcValue = dataRow.DcCategoryValues.Find(x => x.ColumnIndex == columnIndex);
                        if (dcNames[j].ValueType == CategoryValueType.HV)
                        {
                            value = dcValue.Hv.Value;
                        }
                        else if (dcNames[j].ValueType == CategoryValueType.LV)
                        {
                            value = dcValue.Lv.Value;
                        }
                        else
                        {
                            value = dcValue.Nv.Value;
                        }

                        // Rename "MC601 E1 + 10%" to "MC601_E1 + 10%"
                        if (value.Contains("%"))
                        {
                            string pMode = re.Match(value).Groups["pMode"].ToString();
                            string others = re.Match(value).Groups["others"].ToString();
                            value = pMode + "_" + others;
                        }

                        if (i == _specialCategorylst.Count - 1)
                        {
                            if (j == dcNames.Count - 1)
                            {
                                sw.Write(value + "\r\n");
                            }
                            else
                            {
                                sw.Write(value + "\t");
                            }
                        }
                        else
                        {
                            sw.Write(value + "\t");
                        }
                    }
                }
            }

            sw.Flush();
            sw.Close();
        }

        private Dictionary<string, List<DcCategoryName>> GetSpecialColumnNumber(TestSettingData tsData)
        {
            var result = new Dictionary<string, List<DcCategoryName>>();
            List<DcCategoryName> categeoryNames = tsData.DcCategorys;
            foreach (DcCategoryName dcName in categeoryNames)
            {
                string categoryName = dcName.CategoryName.ToUpper();
                if (_specialCategorylst.Contains(categoryName))
                {
                    if (result.ContainsKey(categoryName))
                    {
                        result[categoryName].Add(dcName);
                    }
                    else
                    {
                        result.Add(categoryName, new List<DcCategoryName> { dcName });
                    }
                }
            }
            return result;
        }
    }
}
