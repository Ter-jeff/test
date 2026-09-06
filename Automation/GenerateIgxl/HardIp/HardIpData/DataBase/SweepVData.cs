using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.Utility;

namespace Automation.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class SweepVData
    {
        public string PinName { get; set; }
        public string Type { get; set; }
        public string Start { get; set; }
        public string Stop { get; set; }
        public string Step { get; set; }
        public bool IsNeedConvert { get; set; }
        public string Axis { get; set; }
        public bool IsEquation { get; set; }
        public string Order { get; set; } = "";
        public string VoltageList { get; set; }
        public string InstanceVoltage { get; set; }
        public string ForceType { get; set; }
        public string CustomSetting { get; set; }

        public string Operand
        {
            get
            {
                if (double.TryParse(Start, out double _) && double.TryParse(Stop, out double _))
                {
                    return double.Parse(Start) < double.Parse(Stop) ? "+" : "-";
                }

                IsEquation = true;

                return "+";
            }
        }
        public string Comparator
        {
            get
            {
                if (double.TryParse(Start, out double _) && double.TryParse(Stop, out double _))
                {
                    return double.Parse(Start) <= double.Parse(Stop) ? "<" : ">";
                }

                return "<";
            }
        }
        public bool CheckStep
        {
            get
            {
                if ((Operand == "+" && Step != null && Step.StartsWith("-")) || (Operand == "-" && !Step.StartsWith("-")))
                {
                    return false;
                }

                return true;
            }
        }

        public SweepVData(string loopStr,
            bool isConvert = true, string instanceVoltage = "", string forceType = "", string customSetting = "")
        {
            IsNeedConvert = isConvert;
            IsEquation = false;
            Axis = loopStr.Split(',')[0];
            Type = "";
            InstanceVoltage = instanceVoltage.ToUpper();
            string info = loopStr;
            if (loopStr.Split(',').Length == 4)
            {
                PinName = loopStr.Split(',')[0];
                {
                    Start = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[1], true);
                    Stop = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[2], true);
                    Step = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[3], true);
                }
            }
            ForceType = forceType;
            CustomSetting = customSetting;
        }

        public SweepVData(string sweepStr, string label,
            bool isDcSpecSweep = true, bool isConvert = false, string order = "", string instanceVoltage = "", string forceType = "", string customSetting = "")
        {
            IsNeedConvert = isConvert;
            IsEquation = false;
            Axis = label;
            Order = order;
            Type = isDcSpecSweep ? "" : $"SrcCodeIndex{label}";
            InstanceVoltage = instanceVoltage.ToUpper();
            string info = sweepStr.Split(':').Last();
            if (sweepStr.Split(':').Length == 2 || sweepStr.Split(':').Length == 3)
            {
                if (isDcSpecSweep)
                {
                    PinName = DataConvertor.ConvertValueWithGlbSpec(sweepStr.Split(':')[0], true);
                }
                else
                {
                    PinName = sweepStr.Split(':')[0];
                }

                if (info.Split(',').Length == 3)
                {
                    Start = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[0], true);
                    Stop = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[1], true);
                    IsEquation = !double.TryParse(Start, out double _) && double.TryParse(Stop, out double _);
                    if (IsEquation)
                    {
                        Step = StepEvaluator.Evaluate(info).ToString();
                    }
                    else
                    {
                        Step = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[2], true);
                    }
                }
                else
                {
                    Start = "0";
                    Stop = "0";
                    Step = "0";
                }
            }
            ForceType = forceType;
            CustomSetting = customSetting;
        }

        public SweepVData(string pinName, string voltageList, string label, string order, string instanceVoltage,
            string forceType = "", string customSetting = "")
        {
            PinName = pinName;
            VoltageList = voltageList;
            Axis = label;
            Order = order;
            Type = label;
            InstanceVoltage = instanceVoltage.ToUpper();
            ForceType = forceType;
            CustomSetting = customSetting;
        }

        public SweepVData(SweepVData other)
        {
            if (other == null)
            {
                return;
            }

            PinName = other.PinName;
            Type = other.Type;
            Start = other.Start;
            Stop = other.Stop;
            Step = other.Step;
            IsNeedConvert = other.IsNeedConvert;
            Axis = other.Axis;
            IsEquation = other.IsEquation;
            Order = other.Order;
            VoltageList = other.VoltageList;
            InstanceVoltage = other.InstanceVoltage;
            ForceType = other.ForceType;
            CustomSetting = other.CustomSetting;
        }

        public SweepVData Copy()
        {
            return new SweepVData(this);
        }
    }
}
