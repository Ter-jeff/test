using System.Collections.Generic;

using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{

    public class TimeRow1P4Converter
    {
        protected int MustHaveColumnCnt = 16;

        public bool NeedCompensate(string[] dataArr)
        {
            return dataArr.Length < MustHaveColumnCnt;
        }

        protected string[] DoCompensate(string[] dataArr)
        {

            if (!NeedCompensate(dataArr))
            {
                return dataArr;
            }

            List<string> dataList = [.. dataArr];
            while (dataList.Count < MustHaveColumnCnt)
            {
                dataList.Add("");
            }
            return [.. dataList];
        }

        public virtual TimingRow ConvertTimeRow(string[] dataArr)
        {
            dataArr = DoCompensate(dataArr);

            var row = new TimingRow
            {
                PinGrpName = dataArr[3],
                PinGrpClockPeriod = dataArr[4],
                PinGrpSetup = dataArr[5],
                DataSrc = dataArr[6],
                DataFmt = dataArr[7],
                DriveOn = dataArr[8],
                DriveData = dataArr[9],
                DriveReturn = dataArr[10],
                DriveOff = dataArr[11],
                CompareMode = dataArr[12],
                CompareOpen = dataArr[13],
                CompareClose = dataArr[14],
                EdgeMode = dataArr[15],
                Comment = dataArr.Length > 16 ? dataArr[16] : ""
            };

            return row;
        }
    }
}
