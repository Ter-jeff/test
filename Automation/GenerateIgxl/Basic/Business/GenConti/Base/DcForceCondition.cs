using Automation.Reader;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base
{
    public class DcForceCondition
    {
        public string ForceType { set; get; }
        public string ForceValue { set; get; }

        public DcForceCondition()
        {
            ForceType = "";
            ForceValue = "";
        }

        public DcForceCondition(string type, string value)
        {
            ForceType = type;
            ForceValue = value;
        }

        public DcForceCondition(DcTestContiRow row)
        {
            ForceType = row.TestType == ContiType.OpenShort ? "I" : "V";
            string value = string.Join(";", row.GetForceCondition());
            ForceValue = value;
        }
    }
}
