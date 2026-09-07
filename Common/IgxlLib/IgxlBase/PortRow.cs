using System;
using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public class PortRow : IgxlRow
    {
        public const int ConSettingNumber = 10;
        public const int ConPropertyNumber = 10;

        public string PortName { get; set; } = string.Empty;
        public string ProtocolFamily { get; set; } = string.Empty;
        public string ProtocolType { get; set; } = string.Empty;
        public string ProtocolSettings { get; set; } = string.Empty;
        public List<string> ProtocolSettingValues { get; set; } = [];
        public string FunctionName { get; set; } = string.Empty;
        public string FunctionPin { get; set; } = string.Empty;
        public string FunctionProperties { get; set; } = string.Empty;
        public List<string> FunctionPropertyValues { get; set; } = [];
        public string Comment { get; set; } = string.Empty;

        public void AddSetting(string setting)
        {
            if (FunctionPropertyValues.Count > ConPropertyNumber)
            {
                throw new Exception($"PortMap setting number has exceed the Max number: {ConSettingNumber}");
            }
            ProtocolSettingValues.Add(setting);
        }
    }
}
