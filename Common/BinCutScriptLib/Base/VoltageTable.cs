using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut.Flow;

namespace BinCutScriptLib.Base
{
    public class VoltageTable(InstanceSheet instanceSheet)
    {
        public Dictionary<string, List<string>> Table = new(StringExtensions.IgnoreCase);
        public Dictionary<string, string> NewTestSettingString = [];
        public InstanceSheet TestInstanceSheet = instanceSheet;

        public string GetCompareString(string dcSpecName)
        {
            if (Table.TryGetValue(dcSpecName, out List<string>? value))
            {
                return string.Join(",", value);
            }

            return "";
        }

        private static List<string> GetPinGroupsAllPin(List<string> comparePins)
        {
            List<string> compareAllPins = [];
            foreach (string p in comparePins)
            {
                if (BinCutData.PinMapData!.GroupList.Any(x => x.PinName.EqualsIgnoreCase(p)))
                {
                    List<Pin> pinName = BinCutData.PinMapData.GroupList.First(x => x.PinName.EqualsIgnoreCase(p)).PinList;
                    foreach (Pin pin in pinName)
                    {
                        compareAllPins.Add(pin.PinName);
                    }
                }
                else
                {
                    compareAllPins.Add(p);
                }
            }

            return compareAllPins;
        }

        public string GetCompareStringCs(string dcSpecName, List<string> comparePins)
        {
            List<string> compareAllPins = GetPinGroupsAllPin(comparePins);
            if (Table.TryGetValue(dcSpecName, out List<string>? value))
            {
                if (BinCutData.PinMapData != null)
                {
                    var pins = new List<string>();
                    foreach (string pin in Table[dcSpecName])
                    {
                        string[] arr = pin.Split('=');
                        string pinName = arr[0];
                        if (!compareAllPins.Exists(x => x.EqualsIgnoreCase(pinName)))
                        //!BinCutData.PinMapData.GroupList.Exists(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        pins.Add(pin.Replace("=", ":"));
                    }
                    pins.Sort(PinValueComparer);
                    return string.Join(",", pins);
                }
                return string.Join(",", value);
            }
            return "";
        }

        public string GetCompareString(InstanceBinCut instanceBinCut, string bvName)
        {
            string dcSpecName = instanceBinCut.CurDcSpecCondtion;
            string bvType = instanceBinCut.CurInstanceName.EndsWithIgnoreCase("HBV>") ? "HV" : "LV";
            string mode = BinCutAlgorithmService.GetModeByName(bvName);
            string selector = "";
            string allOthers = "";

            (bool flowControl, string? value) = HandleCompareString(instanceBinCut, bvName, ref dcSpecName, bvType, mode, ref selector, ref allOthers);
            if (!flowControl)
            {
                return value!;
            }
            return GetCompareString(dcSpecName);
        }

        private (bool flowControl, string? value) HandleCompareString(InstanceBinCut instanceBinCut, string bvName, ref string dcSpecName, string bvType, string mode, ref string selector, ref string allOthers)
        {
            if (BinCutConfig.FlagT0TxHotFormat || BinCutConfig.FlagT0TxRoomFormat)
            {
                string t0Txjob = BinCutConfig.FlagT0TxHotFormat ? "T0TX_HOT" : "T0TX_ROOM";
                if (!instanceBinCut.CurInstanceName.Contains("OUTSIDE", StringComparison.OrdinalIgnoreCase))
                {
                    if (BinCutData.BinCutFlowTables.Exists(x => x.JobName.EqualsIgnoreCase(t0Txjob)))
                    {
                        BinCutFlowSheetRow? targetRow = BinCutData.BinCutFlowTables.Find(x => x.JobName.EqualsIgnoreCase(t0Txjob))!.Rows.Find(x => x.PerformanceMode.EqualsIgnoreCase(mode) && x.TableType.ToString().EqualsIgnoreCase(bvType));
                        if (targetRow != null)
                        {
                            allOthers = targetRow.AllOther;
                            selector = allOthers.Split(' ').Last();
                        }
                    }
                }
                else
                {
                    string bvType1 = "post";
                    for (int j = 0; j < BinCutData.PostFlowSheets.Count; j++)
                    {
                        if (BinCutData.PostFlowSheets[j].Exists(x => x.JobName.EqualsIgnoreCase(t0Txjob)))
                        {
                            BinCutFlowSheetRow? targetRow = BinCutData.PostFlowSheets[j].Find(x => x.JobName.EqualsIgnoreCase(t0Txjob))!.Rows.Find(x => bvName.Contains(x.PerformanceMode, StringComparison.OrdinalIgnoreCase) && x.TableType.ToString().EqualsIgnoreCase(bvType1));
                            if (targetRow == null)
                            {
                                continue;
                            }

                            allOthers = targetRow.AllOther;
                            selector = allOthers.Split(' ').Last();
                        }
                    }
                }
                dcSpecName = ConvertAllothersToDCspecs(selector, allOthers);
            }

            if (NewTestSettingString.Count != 0)
            {
                if (NewTestSettingString.TryGetValue(dcSpecName, out string? value))
                {
                    return (flowControl: false, value);
                }
            }

            return (flowControl: true, value: null);
        }

        public string GetCompareStringCs(InstanceBinCut instanceBinCut, string bvName)
        {
            string dcSpecName = instanceBinCut.CurDcSpecCondtion;
            string bvType = instanceBinCut.CurInstanceName.EndsWithIgnoreCase("HBV>") ||
                            instanceBinCut.CurInstanceName.EndsWithIgnoreCase("HBV_Cs>") ||
                            instanceBinCut.CurInstanceName.EndsWithIgnoreCase("HBV_Csharp>") ? "HV" : "LV";
            string mode = BinCutAlgorithmService.GetModeByName(bvName);
            string selector = "";
            string allOthers = "";
            string job = BinCutData.Job.JobName;
            List<string> comparePins = BinCutData.PowerPins;

            if (instanceBinCut.CurInstanceName.Contains("OUTSIDE", StringComparison.OrdinalIgnoreCase) && BinCutData.PostFlowSheets != null)
            {
                foreach (BinCutFlowTables sheet in BinCutData.PostFlowSheets)
                {
                    if (!sheet.Exists(x => x.JobName.EqualsIgnoreCase(job)))
                    {
                        continue;
                    }

                    comparePins = [.. sheet.Find(x => x.JobName.EqualsIgnoreCase(job))!.PowerPins.Keys];
                }
            }
            (bool flowControl, string? value) = HandleCompareString(instanceBinCut, bvName, ref dcSpecName, bvType, mode, ref selector, ref allOthers);
            if (!flowControl)
            {
                return value!;
            }
            return GetCompareStringCs(dcSpecName, comparePins);
        }

        // Replicates Framework 4.8 Windows NLS word-sort: '_' acts as a word separator
        // (effective weight ≈ space/32), so it sorts before digits/letters.
        // Format of each element is "pinname:value"; sort key is the pin name portion.
        private static int PinValueComparer(string a, string b)
        {
            int ca = a.IndexOf(':');
            int cb = b.IndexOf(':');
            string keyA = (ca >= 0 ? a[..ca] : a).Replace('_', ' ');
            string keyB = (cb >= 0 ? b[..cb] : b).Replace('_', ' ');
            int cmp = string.Compare(keyA, keyB, StringComparison.OrdinalIgnoreCase);
            return cmp != 0 ? cmp : string.Compare(a, b, StringComparison.Ordinal);
        }

        public static string ConvertAllothersToDCspecs(string selector, string allOthers)
        {
            if (selector == "LV")
            {
                return allOthers.Replace(" LV", "_MIN");
            }
            else if (selector == "NV")
            {
                return allOthers.Replace(" NV", "_TYP");
            }
            else
            {
                return allOthers.Replace(" HV", "_MAX");
            }
        }
    }
}
