using System.Diagnostics.CodeAnalysis;

namespace Automation.GenerateIgxl.Wireless.DVDC.InputObject
{
    [ExcludeFromCodeCoverage]
    public class WirelessData
    {
        #region Property
        public string TrimFuseName { get; set; } = string.Empty;
        public string TrimTarget { get; set; } = string.Empty;
        public string TrimMeas { get; set; } = string.Empty;
        public string TrimCalcEqn { get; set; } = string.Empty;
        public string TrimType { get; set; } = string.Empty;
        public string RegisterAssignment { get; set; } = string.Empty;
        public bool IsNeedPostBurn { get; set; }
        public bool IsDoMeasure { get; set; }
        #endregion

        public WirelessData()
        {
        }

        public WirelessData(WirelessData data)
        {
            TrimFuseName = data.TrimFuseName;
            TrimTarget = data.TrimTarget;
            TrimMeas = data.TrimMeas;
            TrimCalcEqn = data.TrimCalcEqn;
            TrimType = data.TrimType;
            RegisterAssignment = data.RegisterAssignment;
            IsNeedPostBurn = data.IsNeedPostBurn;
            IsDoMeasure = data.IsDoMeasure;
        }
    }
}
