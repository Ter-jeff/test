using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace RfLib.Dvdc.Reader.CapturePostProcess
{
    public class PostProcessSheet
    {
        #region Field
        public string SetupName = "SetupName";

        public Dictionary<string, int> IndexHeader = [];

        #endregion

        #region Properity
        public string SheetName { set; get; }
        public List<PostProcessSheetRow> RowList { set; get; }
        public List<CPPSetup> Setups = [];
        #endregion

        #region Constructor
        public PostProcessSheet()
        {
            SheetName = "";
            RowList = [];
        }
        #endregion


        public void Read(string filepath)
        {

            string? line;
            var sr = new StreamReader(filepath);
            CPPSetup? setup = null;
            string pattern = "";
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Contains(SetupName))
                {
                    IndexHeader.Clear();
                    int index = 0;
                    setup = new CPPSetup();
                    pattern = "";
                    Setups.Add(setup);
                    foreach (string item in line.Split('\t'))
                    {
                        IndexHeader.Add(item, index);
                        index++;
                    }
                }
                else
                {
                    var row = new PostProcessSheetRow();
                    var postcalc = new PostCalcInfo();
                    row.PostCalcs.Add(postcalc);
                    string[] sgmt = line.Split('\t');
                    foreach (KeyValuePair<string, int> item in IndexHeader)
                    {
                        switch (item.Key)
                        {
                            case "SetupName":
                                row.SetupName = sgmt[item.Value];
                                break;
                            case "PatternName":
                                if (string.IsNullOrEmpty(pattern))
                                {
                                    pattern = sgmt[item.Value];
                                }

                                row.PatternName = pattern;
                                break;
                            case "TestName":
                                row.TestName = sgmt[item.Value];
                                break;
                            case "BitWidth":
                                row.BitWidth = sgmt[item.Value];
                                break;
                            case "StoreName":
                                row.StoreName = sgmt[item.Value];
                                break;
                            case "CalcEquation":
                                row.CalcEquation = sgmt[item.Value];
                                postcalc.CalcEquation = sgmt[item.Value];
                                break;
                            case "CalcTestName":
                                row.CalcTestName = sgmt[item.Value];
                                postcalc.CalcTestName = sgmt[item.Value];
                                break;
                            case "CalcStoreName":
                                row.CalcStoreName = sgmt[item.Value];
                                postcalc.CalcStoreName = sgmt[item.Value];
                                break;
                            case "LowLimit":
                                row.LowLimit = sgmt[item.Value];
                                postcalc.LowLimit = sgmt[item.Value];
                                break;
                            case "HiLimit":
                                row.HiLimit = sgmt[item.Value];
                                postcalc.HiLimit = sgmt[item.Value];
                                break;
                        }
                    }
                    setup!.Datas.Add(row);
                }
            }
            sr.Close();
        }

        public Dictionary<string, DataTable> WriteTableNew()
        {
            var result = new Dictionary<string, DataTable>();
            var groups = RowList.GroupBy(p => p.BlockName).ToDictionary(p => p.Key, p => p.ToList());

            var table = new DataTable();

            MemberInfo[] headers = [.. typeof(PostProcessSheetRow).GetProperties().Where(p => p.Name != "BlockName").Select(p => (MemberInfo)p)];

            for (int i = 1; i <= headers.Length; i++)
            {
                table.Columns.Add(i.ToString(), typeof(string));
            }
            int startCol = 1;
            int startRow = 0;

            foreach (KeyValuePair<string, List<PostProcessSheetRow>> dic in groups)
            {
                DataRow dataRow;

                var cppInfo = dic.Value.GroupBy(x => x.SetupName).ToDictionary(x => x.Key, x => x.ToList());

                //start row set
                foreach (KeyValuePair<string, List<PostProcessSheetRow>> setup in cppInfo)
                {
                    bool isFullSweepTrim = setup.Value.SelectMany(p => p.PostCalcs).Any(p => p.CalcEquation.Contains("Calc_BestCode_CPP"));
                    startCol = 1;

                    //var isFirst = true;
                    dataRow = table.NewRow();
                    table.Rows.Add(dataRow);
                    #region Print Header
                    //Print Header
                    foreach (MemberInfo row in headers)
                    {
                        dataRow[startCol - 1] = row.Name;
                        startCol++;
                    }
                    #endregion
                    #region Print Prescription
                    dataRow = table.NewRow();
                    table.Rows.Add(dataRow);
                    dataRow[PostProcessSheetRow.PatternNameIdx] = setup.Value.FirstOrDefault()!.PatternName;
                    dataRow[PostProcessSheetRow.SetupNameIdx] = setup.Key;
                    DataRow fixedRow;
                    startRow++;
                    for (int i = 3; i < 8; i++)
                    {
                        if (i + startRow < table.Rows.Count)
                        {
                            fixedRow = table.Rows[i + startRow];
                        }
                        else
                        {
                            fixedRow = table.NewRow();
                            table.Rows.Add(fixedRow);
                        }

                        switch (i)
                        {
                            case 3:
                                fixedRow[PostProcessSheetRow.SetupNameIdx] = "SIMULATIONFILE:";
                                break;
                            case 4:
                                //rf => false, FW => true
                                //var flag = dic.Key.Equals("FW");

                                if (setup.Value.SelectMany(x => x.PostCalcs).Select(x => x.CalcEquation).Any(x => !string.IsNullOrEmpty(x)) && !isFullSweepTrim)
                                {
                                    fixedRow[PostProcessSheetRow.SetupNameIdx] = "FORCEFLOWFLAG:true";
                                }
                                else
                                {
                                    fixedRow[PostProcessSheetRow.SetupNameIdx] = "FORCEFLOWFLAG:false";
                                }

                                break;
                            case 5:
                                fixedRow[PostProcessSheetRow.SetupNameIdx] = "CAPTUREDATAPRINT:False";
                                break;
                            case 6:
                                if (setup.Value.SelectMany(row => row.PostCalcs).Any(poca => poca.CalcStoreName.EndsWith("_CapTrimData")) && !isFullSweepTrim)
                                {
                                    fixedRow[PostProcessSheetRow.SetupNameIdx] = "CAPTURETRIM:True";
                                }
                                else
                                {
                                    fixedRow[PostProcessSheetRow.SetupNameIdx] = "CAPTURETRIM:False";
                                }

                                break;
                            case 7:
                                if (setup.Value.Select(row => row.BitWidth).All(bw => bw == "0"))
                                {
                                    fixedRow[PostProcessSheetRow.SetupNameIdx] = "CALONLYFLAG:True";
                                }
                                else
                                {
                                    fixedRow[PostProcessSheetRow.SetupNameIdx] = "CALONLYFLAG:False";
                                }

                                break;
                        }
                    }
                    #endregion
                    int tmpRow = startRow;
                    foreach (PostProcessSheetRow row in setup.Value)
                    {
                        if (startRow < table.Rows.Count)
                        {
                            dataRow = table.Rows[startRow];
                        }
                        else
                        {
                            dataRow = table.NewRow();
                            table.Rows.Add(dataRow);
                        }
                        dataRow[PostProcessSheetRow.TestNameIdx] = row.TestName.ToUpper();
                        dataRow[PostProcessSheetRow.StoreNameIdx] = row.StoreName;
                        if (row.PostCalcs.Count > 0)
                        {
                            foreach (PostCalcInfo postCalc in row.PostCalcs)
                            {
                                if (startRow < table.Rows.Count)
                                {
                                    dataRow = table.Rows[startRow];
                                }
                                else
                                {
                                    dataRow = table.NewRow();
                                    table.Rows.Add(dataRow);
                                }
                                dataRow[PostProcessSheetRow.CalcEquationIdx] = postCalc.CalcEquation;
                                dataRow[PostProcessSheetRow.CalcTestNameIdx] = postCalc.CalcTestName.ToUpper();
                                dataRow[PostProcessSheetRow.CalcStoreNameIdx] = postCalc.CalcStoreName;
                                dataRow[PostProcessSheetRow.BitWidthIdx] = postCalc.Bit != 0 ? postCalc.Bit.ToString() : "";
                                dataRow[PostProcessSheetRow.LowLimitIdx] = postCalc.LowLimit;
                                dataRow[PostProcessSheetRow.HiLimitIdx] = postCalc.HiLimit;
                                startRow++;
                            }
                        }
                        else
                        {
                            startRow++;
                        }
                    }
                    if (tmpRow + 6 > startRow)
                    {
                        startRow = tmpRow + 6;
                    }

                    #region marked out

                    //    if (row.PostCalcs.Count > 0)
                    //    {
                    //        foreach (var postCalc in row.PostCalcs)
                    //        {
                    //            if (startRow < table.Rows.Count)
                    //                dataRow = table.Rows[startRow];
                    //            else
                    //            {
                    //                dataRow = table.NewRow();
                    //                table.Rows.Add(dataRow);
                    //            }
                    //            dataRow[PostProcessSheetRow.CalcEquationIdx] = postCalc.CalcEquation;
                    //            dataRow[PostProcessSheetRow.CalcTestNameIdx] = postCalc.CalcTestName;
                    //            dataRow[PostProcessSheetRow.CalcStoreNameIdx] = postCalc.CalcStoreName;
                    //            dataRow[PostProcessSheetRow.BitWidthIdx] = postCalc.bit;
                    //            dataRow[PostProcessSheetRow.LowLimitIdx] = postCalc.LowLimit;
                    //            dataRow[PostProcessSheetRow.HiLimitIdx] = postCalc.HiLimit;
                    //            startRow++;
                    //        }
                    //        if (row.PostCalcs.Count < 4)
                    //            startRow = startRow + 4 - row.PostCalcs.Count;
                    //    }
                    //    else
                    //    {
                    //        startRow = startRow + 4; 
                    //    }
                    //}

                    #endregion
                }
            }

            result.Add("CPP_ARF", table);

            return result;
        }

        public void Write(string folder)
        {
            //var tables = WriteTable();
            //foreach (var dic in tables)
            //{
            //    var filename = Path.Combine(folder, dic.Key + ".txt");
            //    var headers = typeof(PostProcessSheetRow).GetProperties().Where(p => p.Name != "BlockName").Select(p => (MemberInfo)p).ToArray();
            //    var sw = new StreamWriter(filename);

            //    for (int row = 0; row < dic.Value.Rows.Count; row++)
            //    {
            //        var datas = string.Join("\t", dic.Value.Rows[row].ItemArray);
            //        sw.WriteLine(datas);
            //    }
            //    sw.Close();
            //}

            Dictionary<string, DataTable> tabless = WriteTableNew();
            foreach (KeyValuePair<string, DataTable> dic in tabless)
            {
                string filename = Path.Combine(folder, dic.Key + ".txt");
                MemberInfo[] headers = [.. typeof(PostProcessSheetRow).GetProperties().Where(p => p.Name != "BlockName").Select(p => (MemberInfo)p)];
                var sw = new StreamWriter(filename);

                for (int row = 0; row < dic.Value.Rows.Count; row++)
                {
                    string datas = string.Join("\t", dic.Value.Rows[row].ItemArray);
                    sw.WriteLine(datas);
                }
                sw.Close();
            }
        }
    }

    public partial class CPPSetup
    {
        [GeneratedRegex("SIMULATIONFILE", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("FORCEFLOWFLAG", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("CAPTUREDATAPRINT", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex("CAPTURETRIM", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();

        public List<PostProcessSheetRow> Datas = [];
        public string SIMULATIONFILE
        {
            get { return Datas.First(p => MyRegex().IsMatch(p.SetupName)).SetupName; }
        }
        public bool FORCEFLOWFLAG
        {
            get { return Datas.First(p => MyRegex1().IsMatch(p.SetupName)).SetupName.Split(':')[1].EqualsIgnoreCase("true"); }
        }
        public bool CAPTUREDATAPRINT
        {
            get { return Datas.First(p => MyRegex2().IsMatch(p.SetupName)).SetupName.Split(':')[1].EqualsIgnoreCase("true"); }
        }
        public bool CAPTURETRIM
        {
            get { return Datas.First(p => MyRegex3().IsMatch(p.SetupName)).SetupName.Split(':')[1].EqualsIgnoreCase("true"); }
        }
        public string Pattern
        {
            get { return Datas.First(p => !string.IsNullOrEmpty(p.PatternName)).PatternName; }
        }

    }
}
