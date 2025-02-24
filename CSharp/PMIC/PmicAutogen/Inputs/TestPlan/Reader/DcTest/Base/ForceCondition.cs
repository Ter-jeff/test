﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class ForceCondition
    {
        public ForceCondition()
        {
            ForcePins = new List<ForcePin>();
        }

        public List<ForcePin> ForcePins { get; set; }
    }

    [Serializable]
    public class ForcePin : IEquatable<ForcePin>
    {
        public ForcePin()
        {
            PinName = "";
            ForceType = "";
            ForceValue = "";
            ForceJob = "";
            ForceCnt = 1;
            Type = ForceConditionType.Normal;
        }

        public string PinName { get; set; }
        public string ForceType { get; set; }
        public string ForceValue { get; set; }
        public string ForceJob { get; set; }
        public int ForceCnt { get; set; }
        public ForceConditionType Type { get; set; }

        public bool Equals(ForcePin other)
        {
            return other != null && PinName == other.PinName && ForceType == other.ForceType &&
                   ForceValue == other.ForceValue;
        }
    }

    public enum ForceConditionType
    {
        Normal = 0,
        Others
    }
}