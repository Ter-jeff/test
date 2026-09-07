using System;
using System.Collections.Generic;
using System.Linq;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    /// <summary>
    ///Define the force condition for each Test Instance
    /// </summary>
    /// 
    public class ForceCondition
    {
        public List<ForcePin> ForcePins { get; set; } = new List<ForcePin>();

        public ForceCondition() { }

        public ForceCondition(ForceCondition other)
        {
            if (other == null)
            {
                return;
            }

            ForcePins = other.ForcePins?.Select(pin => pin.Copy()).ToList() ?? new List<ForcePin>();
        }

        public ForceCondition Copy()
        {
            return new ForceCondition(this);
        }
    }

    /// <summary>
    /// Define the force pins, type,value in force condition
    /// </summary>
    /// 
    public class ForcePin : IEquatable<ForcePin>
    {
        public string PinName { get; set; } = "";
        public string ForceType { get; set; } = "";
        public string ForceValue { get; set; } = "";
        public string ForceJob { get; set; } = "";
        public List<string> ForceLabelVoltages { get; set; } = new List<string>();
        public int ForceCnt { get; set; } = 1;
        public ForceConditionType Type { get; set; } = ForceConditionType.Normal;
        public string ForceInterPostMeas { get; set; } = "";
        public bool IsRestore { get; set; } = true; //default is restore

        public bool Equals(ForcePin other)
        {
            return other != null && PinName == other.PinName && ForceType == other.ForceType && ForceValue == other.ForceValue;
        }

        public override int GetHashCode()//Must implement GetHashCode() to use Distinct(). Distinct() will compare HasCode first, if the same then invoke Equals()
        {
            return (PinName + ":" + ForceType).GetHashCode();
        }

        public ForcePin Copy()
        {
            return new ForcePin
            {
                PinName = PinName,
                ForceType = ForceType,
                ForceValue = ForceValue,
                ForceJob = ForceJob,
                ForceCnt = ForceCnt,
                IsRestore = IsRestore,
                Type = Type,
                ForceInterPostMeas = ForceInterPostMeas,
                ForceLabelVoltages = new List<string>(ForceLabelVoltages),
            };
        }
    }

    public enum ForceConditionType
    {
        Normal = 0,
        FvMode,
        FiMode,
        Others
    }
}
