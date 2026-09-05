using System.Collections.Generic;
using System.Globalization;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

namespace Automation.GenerateIgxl.Basic.Business.GenGlobalDc.Business
{
    public class GlbSymbolGenerator
    {
        #region Member Function
        public void DefaultGlbSymbols(GlobalSpecSheet globalSpecSheet)
        {
            DefaultShmooGlb(globalSpecSheet);
        }

        public void PlusGlbSymbols(GlobalSpecSheet globalSpecSheet)
        {
            PlusNwireGlb(globalSpecSheet);

            PlusShmooGlb(globalSpecSheet);
        }

        public void TestSettingPowerPinsShmooGlb(GlobalSpecSheet globalSpecSheet, MultiTestSettingSheetsSingleton allTestSettingData)
        {
            foreach (string powerPin in allTestSettingData.PowerPinList)
            {
                var spec = new GlobalSpec($"Shmoo_{powerPin.ToUpper()}_Glb") { Value = "0.775" };
                globalSpecSheet.AddRow(spec);
            }
        }

        /// <summary>
        /// Add some global spec symbols for nWire
        /// </summary>
        /// <param name="globalSpecSheet"></param>
        public void PlusNwireGlb(GlobalSpecSheet globalSpecSheet)
        {
            GlobalSpec spec;
            List<ProtocolAwarePin> nWirePins = NwireSingleton.Instance().SettingInfo.NwirePins;
            foreach (ProtocolAwarePin nWirePin in nWirePins)
            {


                //Add Output clock Level spec
                spec = new GlobalSpec(nWirePin.CreateOutClkLevelSpecName()) { Value = nWirePin.OutClkVoltage.ToString("G15", CultureInfo.InvariantCulture) };
                globalSpecSheet.AddRow(spec);

                //Add Reference Clock Level spec
                spec = new GlobalSpec(nWirePin.CreateRefClkLevelSpecName()) { Value = nWirePin.RefClkVoltage.ToString("G15", CultureInfo.InvariantCulture) };
                globalSpecSheet.AddRow(spec);

                foreach (EnumEquipment equipment in TestPlanStatic.Equipments)
                {
                    //Add Power seq spec
                    spec = new GlobalSpec(nWirePin.CreatePowerSeqSpecName(equipment)) { Value = nWirePin.PowerUpSeq };
                    globalSpecSheet.AddRow(spec);

                    if (!string.IsNullOrEmpty(nWirePin.PowerDownSeq))
                    {
                        //Add Power down seq spec
                        spec = new GlobalSpec(nWirePin.CreatePowerDownSeqSpecName(equipment)) { Value = nWirePin.PowerDownSeq };
                        globalSpecSheet.AddRow(spec);
                    }
                    //Add Freq spec
                    spec = new GlobalSpec(nWirePin.CreateFreqSpecName(equipment)) { Value = nWirePin.Freq.ToString("G15", CultureInfo.InvariantCulture) };
                    globalSpecSheet.AddRow(spec);
                }
            }

            spec = new GlobalSpec(NwireSingleton.SbcGlbSpecName) { Value = NwireSingleton.Instance().SettingInfo.SupportBoardFreq.ToString("0") };
            globalSpecSheet.AddRow(spec);
            //Add SBC spec
        }

        /// <summary>
        /// <summary>
        /// Add spec symbols for Shmoo
        /// </summary>
        /// <param name="globalSpecSheet"></param>
        private void PlusShmooGlb(GlobalSpecSheet globalSpecSheet)
        {
            var spec = new GlobalSpec(SpecFormat.GenGlbSpecSymbol("Shmoo")) { Value = "-1" };
            globalSpecSheet.AddRow(spec);
        }

        /// <summary>
        /// Add default Shmoo_Glb_GLB symbol for Shmoo
        /// </summary>
        /// <param name="globalSpecSheet"></param>
        private void DefaultShmooGlb(GlobalSpecSheet globalSpecSheet)
        {
            var spec = new GlobalSpec("Shmoo_Glb_GLB") { Value = "0" };
            globalSpecSheet.AddRow(spec);
        }
        #endregion
    }
}
