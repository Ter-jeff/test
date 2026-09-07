using System;
using System.Collections.Generic;

namespace EfuseCheckCmdLib.IgxlLogLib
{
    public class UnitUtit
    {
        private static readonly Lazy<UnitUtit> _lazyInstance = new(() => new UnitUtit());
        public static UnitUtit Instance { get { return _lazyInstance.Value; } }

        private Dictionary<string, int>? _dicUnit;
        public Dictionary<string, int> DicUnit
        {
            get
            {
                if (_dicUnit == null)
                {
                    _dicUnit = [];
                    foreach (char c in "abcdefghijklmnopqrstuvwxyzQAZWSXEDCRFVTGBYHNUJMIKOLP%")
                    {
                        _dicUnit[Convert.ToString(c)] = 0;
                    }
                    _dicUnit["f"] = -15;
                    _dicUnit["p"] = -12;
                    _dicUnit["n"] = -9;
                    _dicUnit["u"] = -6;
                    _dicUnit["m"] = -3;
                    _dicUnit["%"] = -2;
                    _dicUnit["k"] = 3;
                    _dicUnit["K"] = 3;
                    _dicUnit["M"] = 6;
                    _dicUnit["G"] = 9;
                    _dicUnit["T"] = 12;

                }
                return _dicUnit;
            }

        }
        private UnitUtit() { }

    }
}
