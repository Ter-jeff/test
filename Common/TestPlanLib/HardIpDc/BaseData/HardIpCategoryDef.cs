using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.Utility;

namespace TestPlanLib.HardIpDc.BaseData
{
    public partial class HardIpCategoryDef(string categoryName)
    {
        [GeneratedRegex(@"[a-zA-Z]\w+")]
        private static partial Regex MyRegex();

        private List<string> _powerPinList = [];
        public string CategoryName { set; get; } = categoryName;
        public List<HardIpDcRow> DataRows { set; get; } = [];
        public string LevelSheet { set; get; } = "";
        public string DcCategory { set; get; } = "";

        public string GetLevelName()
        {
            if (DataRows.Exists(p => p.PinType.Equals(EnumHardIpDcPinType.IoDiff) || p.PinType.Equals(EnumHardIpDcPinType.IoSingle)))
            {
                return "Levels_" + CategoryName;
            }
            if (DataRows.Exists(p => p.PinType.Equals(EnumHardIpDcPinType.Power) && p.Ifold?.Length != 0))
            {
                return "Levels_" + CategoryName;
            }
            if (DataRows.Exists(p => p.PinType.Equals(EnumHardIpDcPinType.LevelIo) && (p.Iol?.Length != 0 || p.Ioh?.Length != 0 || p.Vcl?.Length != 0 || p.Vch?.Length != 0 || p.DriverMode?.Length != 0)))
            {
                return "Levels_" + CategoryName;
            }
            //"Levels_HardIP";
            return "Levels_" + CategoryName;
        }

        public void SetPowerPinList(List<string> powerPinList)
        {
            _powerPinList = powerPinList;
        }

        public List<HardIpSpecValue> GetSpecValueFromDef(HardIpDcRow hardIpDcRow)
        {
            List<HardIpSpecValue> specValueList = [];
            EnumHardIpDcPinType pintype = hardIpDcRow.PinType;
            const bool hasRatio = true;

            string vmain = FilterDcSpecValue(hardIpDcRow.Nv);
            string ifold = FilterDcSpecValue(hardIpDcRow.Ifold);

            string vil = FilterDcSpecValue(hardIpDcRow.Vil);
            string vih = FilterDcSpecValue(hardIpDcRow.Vih);
            string vol = FilterDcSpecValue(hardIpDcRow.Vol);
            string voh = FilterDcSpecValue(hardIpDcRow.Voh);
            string iol = FilterDcSpecValue(hardIpDcRow.Iol);
            string ioh = FilterDcSpecValue(hardIpDcRow.Ioh);
            string vt = FilterDcSpecValue(hardIpDcRow.Vt);
            string vcl = FilterDcSpecValue(hardIpDcRow.Vcl);
            string vch = FilterDcSpecValue(hardIpDcRow.Vch);

            string vid = FilterDcSpecValue(hardIpDcRow.Vid);
            string vod = FilterDcSpecValue(hardIpDcRow.Vod);
            string vicm = FilterDcSpecValue(hardIpDcRow.Vicm);
            if (pintype.Equals(EnumHardIpDcPinType.Power))
            {
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vmain", vmain, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                if (!string.IsNullOrEmpty(hardIpDcRow.Ifold))
                {
                    specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "iFoldLevel", ifold, false));
                }
            }
            if (pintype.Equals(EnumHardIpDcPinType.Dcvi))
            {
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vps", vmain, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                if (!string.IsNullOrEmpty(hardIpDcRow.Ifold))
                {
                    specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Isc", ifold, false));
                }

                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "tdelay", "0", false));
            }
            if (pintype.Equals(EnumHardIpDcPinType.LevelIo))
            {
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vil", vil, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vih", vih, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vol", vol, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "voh", voh, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "iol", iol, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "ioh", ioh, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vt", vt, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vcl", vcl, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vch", vch, hasRatio, false, hardIpDcRow.HvRatio, hardIpDcRow.LvRatio));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "driverMode", hardIpDcRow.DriverMode, false));
            }
            if (pintype.Equals(EnumHardIpDcPinType.IoSingle))
            {
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vil", vil, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vih", vih, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vol", vol, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "voh", voh, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "voh_alt1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "voh_alt2", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "iol", hardIpDcRow.Iol, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "ioh", hardIpDcRow.Ioh, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vt", vt, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vcl", vcl, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "vch", vch, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "voutLoTyp", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "voutHiTyp", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "driverMode", hardIpDcRow.DriverMode, false));
            }
            if (pintype.Equals(EnumHardIpDcPinType.IoDiff))
            {
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vicm", vicm, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vid", vid, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "dVid0", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "dVid1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "dVicm0", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "dVicm1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vod", vod, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vod_alt1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vod_alt2", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "dVod0", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "dVod1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Iol", iol, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Ioh", ioh, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "VodTyp", vod, false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "VocmTyp", "0", false, true));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vt", vt, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vcl", vcl, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "Vch", vch, false));
                specValueList.Add(new HardIpSpecValue(hardIpDcRow.PinName, "DriverMode", hardIpDcRow.DriverMode, false));
            }

            specValueList.RemoveAll(p => p.Value?.Length == 0);
            return specValueList;
        }

        private string FilterDcSpecValue(string specValue)
        {
            if (!MyRegex().IsMatch(specValue))
            {
                return specValue;
            }
            return MyRegex().Replace(specValue, m =>
            {
                return _powerPinList.Exists(p => p.EqualsIgnoreCase(m.Value))
                    ? SpecFormat.GenDcSpecSymbolAtLevelSheet(m.Value, "", "")
                    : m.Value;
            });
        }
    }

    public class HardIpSpecValue : LevelRow
    {
        public bool FromDefault { set; get; }
        public bool HasRatio { set; get; }
        public string Plus { set; get; }
        public string Minus { set; get; }

        public HardIpSpecValue(string pinName, string parameter, string value, bool hasorNot, bool fromDefault = false, string hv = "", string lv = "", string comment = "")
            : base(pinName, parameter, value, comment)
        {
            PinName = pinName;
            Parameter = parameter;
            Value = value;
            Comment = comment;
            HasRatio = hasorNot;
            FromDefault = fromDefault;
            Plus = hv;
            Minus = lv;

            if (Value?.Length == 0)
            {
                //mark by Raze for removing default if the cell content is blank 2017/06/12
                if (Parameter.EqualsIgnoreCase("driverMode"))
                {
                    Value = "HiZ";
                }
                else if (Parameter.EqualsIgnoreCase("Vcl"))
                {
                    Value = "-1";
                }
                else if (Parameter.EqualsIgnoreCase("Vch"))
                {
                    Value = "6";
                }
                else
                {
                    Value = "0";
                }
                FromDefault = true;
            }
        }

    }
}
