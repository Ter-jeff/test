using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

namespace BinCutScriptLib.Base
{
    public class BvResults : List<BvResult>
    {
        public BvLine? Line;

        public string GetComposeString()
        {
            var list = this.Select(x => $"{x.PinName}=" + $"{Math.Floor(x.Voltage) / 1000:F3}").ToList();
            return string.Join(",", list);
        }

        public string GetComposeStringCs(bool decompose)
        {
            if (BinCutData.PinMapData != null && decompose)
            {
                var pins = new List<string>();
                foreach ((string pinName, double voltage) in this.Select(x => (x.PinName, x.Voltage)))
                {
                    if (BinCutData.PinMapData.IsGroupExist(pinName))
                    {
                        foreach (Pin groupItem in BinCutData.PinMapData.GetGroup(pinName)!.PinList)
                        {
                            pins.Add(groupItem.PinName + ":" + (Math.Floor(voltage) / 1000));
                        }
                    }
                    else
                    {
                        pins.Add(pinName + ":" + (Math.Floor(voltage) / 1000));
                    }
                }
                pins.Sort(PinValueComparer);
                return string.Join(",", pins);
            }

            IOrderedEnumerable<BvResult> sorted = this.OrderBy(x => x.PinName.Replace('_', ' '), StringExtensions.IgnoreCase);
            var list = sorted.Select(x => $"{x.PinName}:" + (x.Voltage / 1000)).ToList();
            return string.Join(",", list);
        }

        private static int PinValueComparer(string a, string b)
        {
            int ca = a.IndexOf(':');
            int cb = b.IndexOf(':');
            string keyA = (ca >= 0 ? a[..ca] : a).Replace('_', ' ');
            string keyB = (cb >= 0 ? b[..cb] : b).Replace('_', ' ');
            int cmp = string.Compare(keyA, keyB, StringComparison.OrdinalIgnoreCase);
            return cmp != 0 ? cmp : string.Compare(a, b, StringComparison.Ordinal);
        }
    }
}
