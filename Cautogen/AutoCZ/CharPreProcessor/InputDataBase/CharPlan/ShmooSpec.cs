using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan
{
    public class ShmooSpec
    {
        public string PowerCondition
        {
            get
            {
                if ((Stop == Start || Stop == "") && Regex.IsMatch(Name.ToLower(), "^vdd.*"))
                {
                    return Start.Trim();
                }

                return "";
            }
        }
        public string Type;
        public int ColIndex { get; set; }
        public string Name { get; set; }
        public bool IsValt { get; set; }
        public bool IsVret { get; set; }
        public string Start { get; set; }
        public string Stop { get; set; }
        public string Step { get; set; }
        public bool IsContainsOrder { get; set; }
        public string Order { get; set; }

        public bool IsSweep
        {
            get
            {
                if (Start != "" && Stop != "" && Start != Stop)
                {
                    return true;
                }

                return Step != "" && Start != Stop;
            }
        }

        public bool IsAllBlank
        {
            get
            {
                return Start == "" && Stop == "" && Step == "";
            }
        }

        public ShmooSpec()
        {
            Start = "";
            Stop = "";
            Step = "";
            //PowerCondition = "";
            IsContainsOrder = false;
            Order = "";
            IsValt = false;
            IsVret = false;
        }

        public ShmooSpec(ShmooSpec item)
        {
            Copy(item);
        }

        public void Copy(ShmooSpec item)
        {
            Name = item.Name;
            Start = item.Start;
            Stop = item.Stop;
            Step = item.Step;
            //PowerCondition = item.PowerCondition;
            IsContainsOrder = item.IsContainsOrder;
            Order = item.Order;
            IsValt = item.IsValt;
            IsVret = item.IsVret;
            Type = item.Type;
        }

        public bool IsPowerPin
        {
            get
            {
                string pinName = Name.ToLower();
                // check if pin in PinMap is power pin
                if (UtilityMain.UtilityData.PinList.TryGetValue(pinName, out PinInfo value))
                {
                    string pinType = value.PinType.ToLower();
                    // check type of pin group (whos pinType in pinMap is "")
                    if (string.IsNullOrEmpty(pinType))
                    {
                        pinType = _GetPinTypeOfPinGroup(Name);
                        if (pinType == null && Regex.IsMatch(Name, "^vdd", RegexOptions.IgnoreCase))
                        {
                            return true; // default assume power pin
                        }

                        if (pinType.Equals("power"))
                        {
                            return true;
                        }
                    }
                    // check type of single pin
                    else if (pinType.Equals("power"))
                    {
                        return true;
                    }
                }

                // check if pin in PinList Sheet is power pin
                // todo : refractory query pingroup list 
                if (UtilityMain.UtilityData.PinGroupList.ContainsKey(pinName) &&
                    !UtilityMain.UtilityData.PinList[UtilityMain.UtilityData.PinGroupList[pinName].First().ToLower()]
                        .PinType.Equals("power", StringComparison.OrdinalIgnoreCase))
                {
                }

                return false;
            }
        }

        private static string _GetPinTypeOfPinGroup(string pinGroupName)
        {
            if (!UtilityMain.UtilityData.PinGroups.TryGetValue(pinGroupName, out List<string> pinGroup))
            {
                return null;
            }

            if (pinGroup == null)
            {
                return null;
            }

            string firstPinName = pinGroup.FirstOrDefault();
            if (firstPinName == null)
            {
                return null;
            }

            PinInfo pin = UtilityMain.UtilityData.PinList[firstPinName.ToLower()];
            return pin == null ? null : pin.PinType;
        }
    }
}
