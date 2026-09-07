using System.Collections.Generic;
using System.Linq;

namespace RfLib.Dvdc.Reader.CapturePostProcess
{
    public class PostProcessSheetRow
    {
        public const int BlockNameIdx = 0;
        public const int SetupNameIdx = 0;
        public const int PatternNameIdx = 1;
        public const int TestNameIdx = 2;
        public const int BitWidthIdx = 3;
        public const int StoreNameIdx = 4;
        public const int CalcEquationIdx = 5;
        public const int CalcTestNameIdx = 6;
        public const int CalcStoreNameIdx = 7;
        public const int LowLimitIdx = 8;
        public const int HiLimitIdx = 9;

        public string BlockName { get; set; }
        public string SetupName { get; set; }
        public string PatternName { get; set; }
        public string TestName { get; set; }
        public string BitWidth { get; set; }
        public string StoreName { get; set; }
        public string CalcEquation { get; set; }
        public string CalcTestName { get; set; }
        public string CalcStoreName { get; set; }
        public string LowLimit { get; set; }
        public string HiLimit { get; set; }

        public List<PostCalcInfo> PostCalcs;

        public PostProcessSheetRow()
        {
            BlockName = "";
            SetupName = "";
            PatternName = "";
            TestName = "";
            BitWidth = "0";
            StoreName = "";
            CalcEquation = "";
            CalcTestName = "";
            CalcStoreName = "";
            LowLimit = "";
            HiLimit = "";
            PostCalcs = [];
        }
        //CalcEquation	CalcTestName	CalcStoreName	LowLimit	HiLimit
        public void AnalyzeCalc(string line, string testname, string lowLimit, string highLimit, string unit, bool isTrimStart, bool isFullSweep)
        {
            //List<Error> errormsgs = new List<Error>();

            int i_name = -1;
            var postCalcItem = new PostCalcInfo
            {
                CalcTestName = testname
            };
            if (string.IsNullOrEmpty(highLimit))
            {
                postCalcItem.HiLimit = "NA";
            }
            else
            {
                postCalcItem.HiLimit = highLimit;
            }

            postCalcItem.HiLimit += unit;

            if (string.IsNullOrEmpty(lowLimit))
            {
                postCalcItem.LowLimit = "NA";
            }
            else
            {
                postCalcItem.LowLimit = lowLimit;
            }

            postCalcItem.LowLimit += unit;
            foreach (string subline in line.Trim(',').Replace("=", ":").Split(';'))
            {
                i_name++;
                string key = subline.Split(':')[0].Trim().ToLower();
                string value = subline.Split(':')[1].Trim();
                string result = new([.. value.Where(c => !char.IsWhiteSpace(c))]);
                switch (key)
                {
                    case "calc":
                        postCalcItem.CalcEquation = result;
                        break;
                    case "storename":
                        StoreName = value;
                        break;
                    case "calcstorename":
                        postCalcItem.CalcStoreName = value;
                        if (isTrimStart && !isFullSweep)
                        {
                            postCalcItem.CalcStoreName += "_CapTrimData";
                        }

                        break;
                }

            }

            string[] csNames = postCalcItem.CalcStoreName.Split(',');
            string[] tNames = testname.Split(',');

            if (tNames.Length > 1 && csNames.Length != tNames.Length)
            {
                ;
            }
            //errormsgs.Add(new Error { 
            //        ErrorType = "CalcError", 
            //        Message = string.Format("CalcStoreName : {0} != TestName : {1}", csNames.Length, tNames.Length), 
            //        ErrorLevel = ErrorLevel.Warning });
            else if (tNames.Length == 1 && csNames.Length == 1)
            {
                PostCalcs.Add(new PostCalcInfo(postCalcItem.CalcEquation, testname, postCalcItem.CalcStoreName, postCalcItem.LowLimit, postCalcItem.HiLimit, postCalcItem.Bit));
            }
            else
            {
                for (int i = 0; i < csNames.Length; i++)
                {
                    string calceqa = i == 0 ? postCalcItem.CalcEquation : "";

                    string[] splitCNName = tNames.Length == 1 ?
                        tNames[0].Split('_') : tNames[i].Split('_');
                    if (tNames.Length == 1)
                    {
                        if (splitCNName.Length >= 3 && !string.IsNullOrEmpty(csNames[i]))
                        {
                            splitCNName[3] = csNames[i].Replace("_", "").ToUpper();
                        }
                    }

                    PostCalcs.Add(new PostCalcInfo(calceqa, string.Join("_", splitCNName), csNames[i], postCalcItem.LowLimit, postCalcItem.HiLimit, i == 0 ? postCalcItem.Bit : 0));
                }
            }
        }
    }
}
