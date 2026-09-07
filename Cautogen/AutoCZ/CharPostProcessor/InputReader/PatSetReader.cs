using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.InputReader
{
    public class PatSetReader
    {
        /* Field */
        private const string PatSetSheetName = "PatSets_All";
        private const string PatSubrSheetName = "Pattern_Subroutine";
        public string BasedPatSetSheet { private get; set; }
        public string BasePatSubrSheet { private get; set; }
        public PatSetSheet BasePatSetSheetAll { get; set; }
        public PatSetSubSheet BasePatSubrSheetAll { get; set; }

        /* Member function */
        public void ReadPatSetTxt()
        {
            BasePatSetSheetAll = new PatSetSheet(PatSetSheetName);
            string[] lines = File.ReadAllLines(BasedPatSetSheet);
            const string lStrHeader = "Pattern Set";

            int lIStartRowNum = 1;
            for (int i = lIStartRowNum; i < lines.Length; i++)
            {
                if (!Regex.IsMatch(lines[i], lStrHeader))
                {
                    continue;
                }

                lIStartRowNum = i + 1;
                break;
            }

            for (int i = lIStartRowNum; i < lines.Length; i++)
            {
                string[] datas = lines[i].Split('\t');
                if (datas.Length < 10)
                {
                    continue;
                }

                BasePatSetSheetAll.AddRow(ReadPatSetRow(datas));
            }
        }

        public void ReadPatSubrTxt()
        {
            BasePatSubrSheetAll = new PatSetSubSheet(PatSubrSheetName);
            string[] lines = File.ReadAllLines(BasePatSubrSheet);

            const int lIStartRowNum = 3; // no header

            for (int i = lIStartRowNum; i < lines.Length; i++)
            {
                string[] datas = lines[i].Split('\t');
                BasePatSubrSheetAll.Rows.Add(new PatSetSubRow
                {
                    PatternFileName = datas[1],
                    Comment = datas.Length >= 3 ? datas[2] : ""
                });
            }
        }

        public static PatSet ReadPatSetRow(IList<string> datas)
        {
            var row = new PatSet { PatSetName = datas[1] };
            row.AddRow(new PatSetRow
            {
                TdGroup = datas[2],
                TimeDomain = datas[3],
                Enable = datas[4],
                File = datas[5],
                Burst = datas[6],
                StartLabel = datas[7],
                StopLabel = datas[8],
                Comment = datas[9],
            });
            return row;
        }
    }
}
