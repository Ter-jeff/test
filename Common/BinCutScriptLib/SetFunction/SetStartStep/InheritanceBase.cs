using BinCutScriptLib.Base;
using BinCutScriptLib.Static;

namespace BinCutScriptLib.SetFunction.SetStartStep
{
    public class InheritanceBase
    {
        protected static void MoveStepByBin(PowerZone powerZone, int bin)
        {
            if (powerZone.PossibleSteps[powerZone.Step].Bin == 1)
            {
                while (powerZone.PossibleSteps[powerZone.Step].Bin != bin)
                {
                    if (powerZone.Step < powerZone.GetPosCount() - 1)
                    {
                        powerZone.Step++;
                        powerZone.FinalStep = powerZone.Step;
                        powerZone.IdsMode = EnumIdsMode.Linear;
                        powerZone.StopStep = powerZone.GetPosCount() - 1;
                        powerZone.Bin = powerZone.PossibleSteps[powerZone.Step].Bin;
                    }
                    else
                    {
                        if (!BinCutConfig.IsDoAll)
                        {
                            powerZone.SearchStatus = EnumSearchStatus.BinOut;
                            //powerZone.FinalStep = -1; //若finalStep=-1表示繼承關係找不到任何更大的電壓, finalStep填為-1
                            powerZone.IdsMode = EnumIdsMode.Linear;
                        }
                        break;
                    }
                }
            }
            else
            {
                for (int i = 0; i < powerZone.PossibleSteps.Count; i++)
                {
                    if (powerZone.PossibleSteps[i].Bin == bin)
                    {
                        powerZone.Step = i;
                        powerZone.FinalStep = powerZone.Step;
                        powerZone.IdsMode = EnumIdsMode.Linear;
                        powerZone.StopStep = powerZone.GetPosCount() - 1;
                        powerZone.Bin = powerZone.PossibleSteps[powerZone.Step].Bin;
                    }
                }
            }
        }
    }
}
