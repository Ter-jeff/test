using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

namespace Automation.GenerateIgxl.PreAction.AddPinGrp
{
    public class SpecialPinGrpPlus
    {
        private const string CorePowerGrpName = "CorePower";

        public void WorkFlow(MultiTestSettingSheetsSingleton testSettingSheetsSingleton)
        {
            if (testSettingSheetsSingleton.VrsPowerPinList.Count == 0)
            {
                return;
            }

            PinGroup allPowerPinGrp = new PinGroup(CorePowerGrpName, PinMapConst.TypePower);
            foreach (string pin in testSettingSheetsSingleton.VrsPowerPinList)
            {
                allPowerPinGrp.AddPin(pin);
            }

            AddPinGrp(allPowerPinGrp);
        }

        private void AddPinGrp(PinGroup pinGroup)
        {
            if (pinGroup.PinList.Count > 0)
            {
                TestProgram.IgxlWorkBk.PinMapPair.Value?.AddRow(pinGroup);
            }
        }

    }
}
