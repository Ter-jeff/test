using System;
using System.Collections.Generic;
using System.Linq;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class SweepCode
    {
        public enum SweepType
        {
            Common,
            Custom
        }

        private const string BinToGray = "BinToGray";
        private const string SignBinToGray = "BinToGray_Sign";
        public string SendBitName { set; get; } = string.Empty;
        public string SweepInfo { set; get; } = string.Empty;
        public string Algorithm { set; get; } = string.Empty;
        public int Width { set; get; }

        public int Start
        {
            get
            {
                if (Type == SweepType.Custom)
                {
                    return 0;
                }

                int outVal = 0;
                if (SweepInfo.Split(',').Any())
                {
                    int.TryParse(SweepInfo.Split(',')[0], out outVal);
                }

                return outVal;
            }
        }
        public int End
        {
            get
            {
                int outVal = 0;
                if (SweepInfo.Split(',').Length > 1)
                {
                    int.TryParse(SweepInfo.Split(',')[1], out outVal);
                }

                return outVal;
            }
        }
        public int Step
        {
            get
            {
                if (Type == SweepType.Custom)
                {
                    return SweepInfo.Split(',').Length;
                }

                int outVal = 0;
                if (SweepInfo.Split(',').Length > 2)
                {
                    if (!int.TryParse(SweepInfo.Split(',')[2], out outVal))
                    {
                        return 1;
                    }
                }

                return outVal;
            }
        }
        public string Order { set; get; } = string.Empty;
        public string StepMisc
        {
            get
            {
                if (Type == SweepType.Custom)
                {
                    return "";
                }

                if (SweepInfo.Split(',').Length == 3)
                {
                    if (!int.TryParse(SweepInfo.Split(',')[2], out int _))
                    {
                        return SweepInfo.Split(',')[2];
                    }
                }
                return "";
            }
        }
        public string Misc
        {
            get
            {
                if (Type == SweepType.Custom)
                {
                    return "";
                }

                if (SweepInfo.Split(',').Length == 4)
                {
                    return SweepInfo.Split(',')[3];
                }
                return "";
            }
        }
        public SweepType Type { set; get; } = SweepType.Common;

        public SweepCode()
        {
        }

        public SweepCode(SweepCode other)
        {
            if (other == null)
            {
                return;
            }

            SendBitName = other.SendBitName;
            SweepInfo = other.SweepInfo;
            Algorithm = other.Algorithm;
            Width = other.Width;
            Order = other.Order;
            Type = other.Type;
        }

        public SweepCode Copy()
        {
            return new SweepCode(this);
        }

        public string GetFlowForLoopInfo(bool isCsUsing = false)
        {
            var result = new List<string>();
            if (Type == SweepType.Common)
            {
                result.Add(SendBitName);
                result.Add(Width.ToString());
                if (!isCsUsing)
                {
                    result.Add("0");
                    result.Add("1");
                }
                if (!string.IsNullOrEmpty(Misc))
                {
                    if (Misc.Equals("graycode", StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(BinToGray);
                    }
                    else if (Misc.Equals("graycode_sign", StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(SignBinToGray);
                    }
                    else
                    {
                        result.Add(Misc);
                    }
                }
            }
            else
            {
                result.Add($"{SendBitName}:[{SweepInfo}]:{Width};");
            }
            return string.Join(":", result);
        }
    }
}
