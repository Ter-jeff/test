using System;

namespace BinCutScriptLib.Base
{
    public static class PowerZoneHelpers
    {
        public static double CalculateFinalValue(PowerZone powerZone)
        {
            if (powerZone.FinalStep == -1)
            {
                throw new Exception($"FinalStep of {powerZone.Mode} can not be -1 !!!");
            }

            return powerZone.PossibleSteps[powerZone.FinalStep].BinningProduct != 0 ? powerZone.PossibleSteps[powerZone.FinalStep].BinningProduct : powerZone.PossibleSteps[powerZone.FinalStep].ProductValue;
        }
    }
}
