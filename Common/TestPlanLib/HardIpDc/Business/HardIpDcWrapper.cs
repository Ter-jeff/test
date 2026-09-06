using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;

namespace TestPlanLib.HardIpDc.Business
{
    public partial class HardIpDcWrapper
    {
        [GeneratedRegex(@"_?DIFF\d?$|_DIFF_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();
        #region Field
        private readonly HardIpDcSheet _dcSheet;
        private readonly List<string> _powerPinList;
        private readonly List<string> _ioPinList;
        #endregion

        public HardIpDcWrapper(HardIpDcSheet hardIpDcSheet, List<string> powerPinList, IoInfoSheet ioInfoSheet)
        {
            _dcSheet = hardIpDcSheet;
            _powerPinList = powerPinList;
            if (ioInfoSheet != null)
            {
                _ioPinList = [.. ioInfoSheet.GetBlockPins("Common")];
            }
            else
            {
                _ioPinList = [];
            }
        }

        public HardIpDcSheet? WrapHardIpDc()
        {
            if (_dcSheet == null)
            {
                return null;
            }

            foreach (HardIpCategoryDef defRow in _dcSheet.Rows)
            {
                defRow.SetPowerPinList(_powerPinList);
                foreach (HardIpDcRow row in defRow.DataRows)
                {
                    row.PinType = GetPinType(row.PinName);
                }
            }
            return _dcSheet;
        }

        private EnumHardIpDcPinType GetPinType(string pinName)
        {
            if (_powerPinList.Contains(pinName.ToUpper()))
            {
                return EnumHardIpDcPinType.Power;
            }
            if (_ioPinList.Contains(pinName.ToUpper()))
            {
                return EnumHardIpDcPinType.LevelIo;
            }
            return _regex.IsMatch(pinName)
                && !pinName.EndsWithIgnoreCase("_Port")
                ? EnumHardIpDcPinType.IoDiff
                : EnumHardIpDcPinType.IoSingle;
        }
    }
}
