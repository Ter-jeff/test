using System;
using System.Collections.Generic;

using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.Singleton
{
    public class CharSetupSingleton
    {
        private readonly Dictionary<string, string> _shmooParameterTypeDictionary;

        private static CharSetupSingleton _instance;

        private CharSetupSingleton()
        {
            _shmooParameterTypeDictionary = GetShmooParameterTypeDictionary();
        }

        public static CharSetupSingleton Instance()
        {
            return _instance ?? (_instance = new CharSetupSingleton());
        }

        public Dictionary<string, string> GetShmooParameterTypeDictionary()
        {
            Dictionary<string, string> newList = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (TestProgram.IgxlWorkBk == null)
            {
                return newList;
            }

            #region AC spec
            foreach (KeyValuePair<string, AcSpecSheet> item in TestProgram.IgxlWorkBk.AcSpecSheets)
            {
                AcSpecSheet data = TestProgram.IgxlWorkBk.AcSpecSheets[item.Key];
                foreach (AcSpec row in data.Rows)
                {
                    if (string.IsNullOrEmpty(row.Symbol))
                    {
                        break;
                    }

                    if (!newList.ContainsKey(row.Symbol))
                    {
                        newList.Add(row.Symbol, "AC spec");
                    }
                }
            }
            #endregion

            #region DC spec
            foreach (KeyValuePair<string, DcSpecSheet> item in TestProgram.IgxlWorkBk.DcSpecSheets)
            {
                DcSpecSheet data = TestProgram.IgxlWorkBk.DcSpecSheets[item.Key];
                foreach (DcSpec row in data.Rows)
                {
                    if (!newList.ContainsKey(row.Symbol))
                    {
                        newList.Add(row.Symbol, "DC spec");
                    }
                }
            }
            #endregion

            #region Global Spec
            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value != null)
            {
                List<GlobalSpec> lGlobalSpecsListSource = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.Rows;
                for (int i = 0; i < lGlobalSpecsListSource.Count; i++)
                {
                    GlobalSpec glbSpec = lGlobalSpecsListSource[i];
                    if (!newList.ContainsKey(glbSpec.Symbol))
                    {
                        newList.Add(glbSpec.Symbol, "Global Spec");
                    }
                }
            }
            #endregion

            #region Level
            foreach (KeyValuePair<string, LevelSheet> item in TestProgram.IgxlWorkBk.LevelSheets)
            {
                LevelSheet data = TestProgram.IgxlWorkBk.LevelSheets[item.Key];
                foreach (LevelRow row in data.Rows)
                {
                    if (!newList.ContainsKey(row.Parameter))
                    {
                        newList.Add(row.Parameter, "Level");
                    }
                }
            }
            #endregion

            #region TimeSet
            foreach (KeyValuePair<string, TimeSetBasicSheet> item in TestProgram.IgxlWorkBk.TimeSetSheets)
            {
                TimeSetBasicSheet data = TestProgram.IgxlWorkBk.TimeSetSheets[item.Key];
                foreach (TSet row in data.Rows)
                {
                    if (!newList.ContainsKey(row.Name))
                    {
                        newList.Add(row.Name, "Period");
                    }
                }
            }
            #endregion

            Dictionary<string, string> defaultList = new Dictionary<string, string>
            {
                //Edge: One of the timing edges (Close,Data,Off,On,Open,Ref Offset,Return) 
                {"Close", "Edge"},{"Data", "Edge"},{"Off", "Edge"},{"On", "Edge"},{"Open", "Edge"},{"RefOffset", "Edge"},{"Return", "Edge"},
                {"d0", "Edge"},{"d1", "Edge"},{"d2", "Edge"},{"d3", "Edge"},{"r0", "Edge"},{"r1", "Edge"},
                //Protocol Aware: One of these values: Clock Offset, Drive Delay, Receive Delay, HiZ Delay, Reference Offset. 
                {"ClockOffset", "Protocol Aware"},{"DriveDelay", "Protocol Aware"},{"ReceiveDelay", "Protocol Aware"},{"HiZDelay", "Protocol Aware"},{"ReferenceOffset", "Protocol Aware"}
                //Serial Timing: Timing set defined on a Serial Timing sheet (Master Period,DUT Period,Drive Delay,Receive Delay)
                //{"Master Period", "Serial Timing"},{"DUT Period", "Serial Timing"},{"Drive Delay", "Serial Timing"},{"Receive Delay", "Serial Timing"}
            };

            foreach (KeyValuePair<string, string> item in defaultList)
            {
                if (!newList.ContainsKey(item.Key))
                {
                    newList.Add(item.Key, item.Value);
                }
            }
            return newList;
        }

        public string GetShmooParameterType(string name)
        {
            if (_shmooParameterTypeDictionary.TryGetValue(name, out string type))
            {
                return type;
            }
            return "";
        }

        public static void Initialize()
        {
            _instance = null;
        }

    }
}
