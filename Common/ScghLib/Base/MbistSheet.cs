using ScghLib.Enums;

namespace ScghLib.Base
{
    public class MbistSheet
    {
        public string SheetName { get; set; } = string.Empty;
        public string Jobs { get; set; } = string.Empty;
        public bool IsMultipleSheet { get; set; }
        public bool IsMultiChipLet { get; set; }
        public MbistPatSetType MbistPatSetType { get; set; } = MbistPatSetType.Single;
        public MbistBinTableType MbistBinTableType { get; set; }
        public bool IsCof { get; set; }
        public bool MbistLoop { get; set; } = false;

        public MbistSheet() { }

        public MbistSheet(MbistSheet mbistSheet)
        {
            if (mbistSheet == null)
            {
                return;
            }

            SheetName = mbistSheet.SheetName;
            Jobs = mbistSheet.Jobs;
            IsMultipleSheet = mbistSheet.IsMultipleSheet;
            IsMultiChipLet = mbistSheet.IsMultiChipLet;
            MbistPatSetType = mbistSheet.MbistPatSetType;
            MbistBinTableType = mbistSheet.MbistBinTableType;
            IsCof = mbistSheet.IsCof;
            MbistLoop = mbistSheet.MbistLoop;
        }

        public MbistSheet Copy()
        {
            return new MbistSheet(this);
        }
    }
}
