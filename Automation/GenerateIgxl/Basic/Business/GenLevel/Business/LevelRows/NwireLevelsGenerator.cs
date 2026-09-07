using System;
using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows
{
    public class NwireLevelsGenerator : LevelsGenerator
    {
        protected const string VolVohRatio = "0.25";
        protected const string VtRatio = "0";
        public override void UpdateLevelSheet(ref LevelSheet levelsheet, LevelData levelData = null)
        {
            //Add nWire pin level
            List<ProtocolAwarePin> nWirePins = NwireSingleton.Instance().SettingInfo.NwirePins;
            List<EnumEquipment> equipments = TestPlanStatic.Equipments;
            HashSet<string> nWirePinRefClk = new(StringComparer.OrdinalIgnoreCase);
            foreach (EnumEquipment equipment in equipments)
            {
                foreach (ProtocolAwarePin nWirePin in nWirePins)
                {
                    if (nWirePin.PinType == IoPinType.Diff)
                    {
                        //Differential pin
                        levelsheet.AddDiffLevel(CreateNwireDiffLevel(nWirePin, equipment));
                    }
                    else
                    {
                        //single-end pin
                        levelsheet.AddIoPinLevel(CreateNwireSingleLevel(nWirePin, equipment));
                    }
                    if (levelsheet.Name.Equals("Levels_nWire", StringComparison.OrdinalIgnoreCase) && equipment != EnumEquipment.UltraFlexPlus)
                    {
                        if (nWirePinRefClk.Add(nWirePin.RefClk))
                        {
                            levelsheet.AddIoPinLevel(CreateNwireRefClockPinLevel(nWirePin));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Create Levels for nWire Differential pin
        /// </summary>
        /// <param name="pin">nWire Pin</param>
        /// <returns></returns>
        private DiffLevel CreateNwireDiffLevel(ProtocolAwarePin pin, EnumEquipment equipment)
        {
            string dVid0 = "0";
            string dVid1 = "0";
            string dVicm0 = "0";
            string dVicm1 = "0";
            string vod = "0";
            string vodAlt1 = "0";
            string vodAlt2 = "0";
            string dVod0 = "0";
            string dVod1 = "0";
            string iol = "0";
            string ioh = "0";
            string vodTyp = "0";
            string vocmTyp = "0";
            string vt = "0";

            string pinName = pin.CreateDiffPinGroupName(equipment);
            string vicm = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName() + "*0.5";
            string vid = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName();
            string vcl = "_Vcl_default";
            string vch = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName() + "*1.1";
            string driverMode = "VT";
            DiffLevel level = new DiffLevel(pinName, vicm, vid, dVid0, dVid1, dVicm0, dVicm1, vod, vodAlt1, vodAlt2, dVod0, dVod1, iol, ioh, vodTyp, vocmTyp, vt, vcl, vch, driverMode);
            return level;
        }

        /// <summary>
        /// Create Levels for nWire Single-end Pin
        /// </summary>
        /// <param name="pin"></param>
        /// <returns></returns>
        private IoLevel CreateNwireSingleLevel(ProtocolAwarePin pin, EnumEquipment equipment)
        {
            string vil = "0";
            string vohAlt1 = "0";
            string vohAlt2 = "0";
            string iol = "0";
            string ioh = "0";
            string voutLoTyp = "0";
            string voutHiTyp = "0";

            string pinName = pin.CreatePaClkPinName(equipment);
            string vih = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName();
            string vol = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName() + "*0.5";
            string voh = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName() + "*0.5";
            string vt = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName() + "*0.5";
            string vcl = "_Vcl_default";
            string vch = SpecFormat.SpecValuePrefix + pin.CreateOutClkLevelSpecName() + "*1.1";
            string driverMode = "VT";

            IoLevel level = new IoLevel(pinName, vil, vih, vol, voh, vohAlt1, vohAlt2, iol, ioh, vt, vcl, vch, voutLoTyp, voutHiTyp, driverMode);
            return level;
        }

        private IoLevel CreateNwireRefClockPinLevel(ProtocolAwarePin pin)
        {
            string vil = "0";
            string vohAlt1 = "0";
            string vohAlt2 = "0";
            string iol = "0";
            string ioh = "0";
            string voutLoTyp = "0";
            string voutHiTyp = "0";

            string pinName = pin.RefClk;
            string vih = SpecFormat.SpecValuePrefix + pin.CreateRefClkLevelSpecName();
            string vol = SpecFormat.SpecValuePrefix + pin.CreateRefClkLevelSpecName() + "*" + VolVohRatio;
            string voh = SpecFormat.SpecValuePrefix + pin.CreateRefClkLevelSpecName() + "*" + VolVohRatio;
            string vt = SpecFormat.SpecValuePrefix + pin.CreateRefClkLevelSpecName() + "*" + VtRatio;
            string vcl = "_Vcl_default";
            string vch = SpecFormat.SpecValuePrefix + pin.CreateRefClkLevelSpecName() + "*1.1";
            string driverMode = "VT";
            IoLevel level = new IoLevel(pinName, vil, vih, vol, voh, vohAlt1, vohAlt2, iol, ioh, vt, vcl,
                vch, voutLoTyp, voutHiTyp, driverMode);
            return level;
        }
    }
}
