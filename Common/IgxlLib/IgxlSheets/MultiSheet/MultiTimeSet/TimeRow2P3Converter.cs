using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{
    public class TimeRow2P3Converter : TimeRow1P4Converter
    {
        public TimeRow2P3Converter()
        {
            MustHaveColumnCnt = 17;
        }

        public override TimingRow ConvertTimeRow(string[] dataArr)
        {
            dataArr = DoCompensate(dataArr);
            TimingRow row = base.ConvertTimeRow(dataArr);

            row.CompareRefOffset = dataArr[15];
            row.EdgeMode = dataArr[16];
            row.Comment = dataArr.Length > 17 ? dataArr[17] : "";

            return row;
        }
    }
}
