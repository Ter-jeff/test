using System.Collections.Generic;
using System.IO;

using TestPlanLib.BinCut.BinCutInstance;

namespace TestPlanLib.BinCut
{
    public class BinCutPatternReport
    {
        #region Field
        public List<BinCutPatternRow> Rows = [];
        #endregion

        public void WriteTxt(string outPath, List<string> appendList)
        {
            using StreamWriter sw = File.CreateText(outPath);
            const string header = "Voltage" + "\t" + "Job" + "\t" + "BinningDomain" + "\t" + "PerformanceMode" + "\t" + "Type" + "\t" + "Condition" + "\t" + "InstanceName" + "\t" + "SheetName" + "\t" + "RowNumber" + "\t" + "Instance" + "\t" + "PattrenSetName" +
                                  "\t" + "Pattern" + "\t" + "PatternVer" + "\t" + "InputFileType" + "\t" + "FileName";
            sw.WriteLine(header);

            for (int index = 0; index < Rows.Count; index++)
            {
                BinCutPatternRow row = Rows[index];
                string text = row.Voltage + "\t" + string.Join(",", row.Jobs) + "\t" + row.BinningDomain + "\t" +
                              row.PerformanceMode + "\t" + row.Type + "\t" + row.Condition + "\t" + row.InstanceName +
                              "\t" + row.BinCutInstanceRow!.SheetName + "\t" + row.BinCutInstanceRow.RowNum + "\t" +
                              row.Instance + "\t" +
                              row.PattrenSetName + "\t" + row.Pattern + "\t" + row.PatternVer;
                if (index < appendList.Count)
                {
                    text += "\t" + appendList[index];
                }

                sw.WriteLine(text);
            }
            sw.WriteLine();
        }
    }

    public class BinCutPatternRow(string flowName)
    {
        public string FlowName = flowName;
        public List<string> Jobs = [];
        public string Voltage = "";
        public string BinningDomain = "";
        public string PerformanceMode = "";
        public string Type = "";
        public string Condition = "";
        public string InstanceName = "";
        public string PattrenSetName = "";
        public string Pattern = "";
        public string PatternVer = "";
        public string Enbable = "";
        public string DeviceCondition = "";
        public string Env = "";
        public string Instance = "";
        public BinCutInstanceRow? BinCutInstanceRow { get; set; }
        public int RowNum;
        public string SheetName = "";

        public bool IsInterpolateSkip;
    }
}
