using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{
    [DebuggerDisplay("{Name}")]
    public class ComTimeSetBasicSheet : TimeSetBasicSheet
    {
        public enum EnumBlockType
        {
            Common,
            Efuse,
            Scan,
            Mbist,
            HardIp,
            BScan,
            None
        }

        public struct TSetEqnVarMap
        {
            public string TSetName;
            public Dictionary<string, double> DictVariable;
        }

        public const string HardIp = "HardIP";

        public override string SheetType => "DTTimesetBasicSheet";

        private readonly List<string> _shiftInTSet = [];

        public bool IsMultiShiftInTSet
        {
            get
            {
                return _shiftInTSet.Count > 1;
            }
        }

        public string GetMultiShiftInStr
        {
            get { return string.Join(",", _shiftInTSet); }

        }

        public List<TSetEqnVarMap> AllTSetEqnVariable
        {
            get
            {
                var tSetEqnVariable = new List<TSetEqnVarMap>();
                var mainCommentVariable = new Dictionary<string, double>();
                foreach (TSet tSet in Rows)
                {
                    if (tSet is not ComTimeSetBasic comTsb)
                    {
                        throw new Exception("");
                    }

                    foreach (KeyValuePair<string, double> pair in comTsb.SubCommentVariable)
                    {
                        if (!mainCommentVariable.ContainsKey(pair.Key))
                        {
                            mainCommentVariable.Add(pair.Key, pair.Value);
                        }
                    }

                    TSetEqnVarMap eqnVarMapObj;
                    eqnVarMapObj.TSetName = comTsb.Name;
                    eqnVarMapObj.DictVariable = new Dictionary<string, double>(mainCommentVariable);
                    tSetEqnVariable.Add(eqnVarMapObj);
                }
                return tSetEqnVariable;
            }
        }

        public ComTimeSetBasicSheet(string sheetName) : base(sheetName)
        {
        }

        public ComTimeSetBasicSheet(string sheetName, string timingMode) : base(sheetName, timingMode)
        {
        }

        public ComTimeSetBasicSheet(string sheetName, string timingMode, string masterTimeSet, string timeDomain, string strobeRefSetup) :
            base(sheetName, timingMode, masterTimeSet, timeDomain, strobeRefSetup)
        {
        }

        public void AddShiftInTSetName(string tSet)
        {
            _shiftInTSet.Add(tSet);
        }

        public void InsertAlarmDataInFirstRow(string alarmString)
        {
            var alarmTSet = new ComTimeSetBasic { Name = alarmString + " Please check it" };
            alarmTSet.AddTimingRow(new TimingRow());
            alarmTSet.AddTimingRow(new TimingRow());
            Rows.Insert(0, alarmTSet);
        }

        public static string GetTimeSetCategory(string sheetName)
        {
            string[] items = sheetName.Split(['_'], StringSplitOptions.RemoveEmptyEntries);
            //Soc/Gfx/Cpu/HardIp/Other...... default set to HardIp
            string moduleBlock = HardIp;
            if (items.Length >= 5)
            {
                //Main block part
                const string moduleCpu = "Cpu";
                const string moduleGfx = "Gfx";
                const string moduleSoc = "Soc";
                switch (items[2].ToUpper()[0])
                {
                    case 'S':
                    case 'H':
                        moduleBlock = moduleSoc;
                        break;
                    case 'C':
                        moduleBlock = moduleCpu;
                        break;
                    case 'L':
                        moduleBlock = moduleGfx;
                        break;
                }
                //Scan/Mbist/Other

                EnumBlockType blockType = GetBlockType(items);
                if (blockType == EnumBlockType.None)
                {
                    return "AC_" + sheetName;
                }

                if (blockType == EnumBlockType.HardIp)
                {
                    return HardIp;
                }

                if (moduleBlock == HardIp || blockType.ToString() == HardIp)
                {
                    return HardIp;
                }

                return moduleBlock + blockType;
            }
            return HardIp;
        }

        public static EnumBlockType GetBlockType(string[] items)
        {
            if (items.ToList().Exists(x => x.EqualsIgnoreCase("Scan") ||
                x.EqualsIgnoreCase("Saa")))
            {
                return EnumBlockType.Scan;
            }

            string block = items[3];
            if (block.EqualsIgnoreCase("AN"))
            {
                return EnumBlockType.HardIp;
            }

            if (block.EqualsIgnoreCase("JT"))
            {
                return EnumBlockType.HardIp;
            }

            if (block.EqualsIgnoreCase("SC"))
            {
                return EnumBlockType.Scan;
            }

            if (block.EqualsIgnoreCase("BI"))
            {
                return EnumBlockType.Mbist;
            }

            if (items.ToList().Exists(x => x.EqualsIgnoreCase("bsr") ||
                                          x.EqualsIgnoreCase("mbist")))
            {
                return EnumBlockType.Mbist;
            }

            return EnumBlockType.None;
        }
    }
}
