using System.Collections.Generic;

using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{
    public class ComTimeSetBasic : TSet
    {
        public Dictionary<string, double> SubCommentVariable = [];
        public List<string> SubContextVariable = [];
        public Dictionary<string, double> ShiftInReserve = [];
    }
}
