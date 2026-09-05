using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.Relay.RelayConst;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class NwireInstanceGenerator
    {

        /// <summary>
        /// Add HardIP nWire instances to "TestInst_Common" sheet
        /// </summary>
        /// <param name="planDic"></param>
        public void GenNwireInstance(Dictionary<string, HardIpSheet> planDic, List<InstanceSheet> instSheets)
        {
            InstanceSheet commonInstanceSheet;
            bool exist = false;
            string key = "";
            foreach (KeyValuePair<string, InstanceSheet> insSheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (string.Equals(insSheet.Value.Name, OutputConst.RelayInstName, StringComparison.OrdinalIgnoreCase))
                {
                    exist = true;
                    key = insSheet.Key;
                    break;
                }
            }

            if (!exist)
            {
                commonInstanceSheet = new InstanceSheet(OutputConst.RelayInstName);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, commonInstanceSheet);
            }
            else
            {
                commonInstanceSheet = TestProgram.IgxlWorkBk.InsSheets[key];
            }
            //Record all relay instances names to prevent duplicate relay instance
            var allNwireInstance = commonInstanceSheet.Rows.Where(a => a.TestName.ContainsIgnoreCase(NwireInstance.FreeRunClkEnable.ToLower())).Select(a => a.TestName).ToList();
            foreach (string sheet in planDic.Keys)
            {
                foreach (HardIpPattern pattern in planDic[sheet].Rows)
                {
                    List<Timing> timings = pattern.GetTimingsByAc().FindAll(s => NwireSingleton.Instance()
                        .SettingInfo.NwirePins.Find(a => a.CreatePinNameWithDiff().Equals(s.Name, StringComparison.InvariantCulture)) != null);
                    if (timings.Count > 0)
                    {
                        List<InstanceRow> nWireInstance = GenerateNwire(timings, pattern, instSheets);
                        foreach (InstanceRow row in nWireInstance)
                        {
                            if (!allNwireInstance.Contains(row.TestName, StringComparer.OrdinalIgnoreCase))
                            {
                                commonInstanceSheet.AddRow(row);
                                allNwireInstance.Add(row.TestName);
                            }
                        }
                    }
                }
            }
        }

        private List<InstanceRow> GenerateNwire(List<Timing> timings, HardIpPattern pattern, List<InstanceSheet> instSheets)
        {
            var result = new List<InstanceRow>();
            List<InstanceRow> instsheetAll = instSheets.SelectMany(x => x.Rows).ToList();
            string timeSets;
            if (instsheetAll.Exists(y => y.Args[0].Equals(pattern.Pattern.GetLastPayload(), StringComparison.OrdinalIgnoreCase)))
            {
                timeSets = instsheetAll.FirstOrDefault(y => y.Args[0].Equals(pattern.Pattern.GetLastPayload(), StringComparison.OrdinalIgnoreCase)).TimeSets;
            }
            else
            {
                timeSets = "";
            }

            string blockName = pattern.SheetName.ToUpper().Replace("HARDIP_", "").Replace(" ", "");
            string acCategory = GetAcCategoryByAc(pattern, blockName, timeSets);
            foreach (ProtocolAwarePin defalutNwire in NwireSingleton.Instance().SettingInfo.NwirePins)
            {
                var nWireRow = new InstanceRow();
                Timing timing = timings.Find(a => a.Name.ToLower().Equals(defalutNwire.CreatePinNameWithDiff(), StringComparison.OrdinalIgnoreCase));
                if (timing == null)
                {
                    nWireRow.TestName = NwireInstance.FreeRunClkEnable + "_" + defalutNwire.CreatePinNameWithDiff() + "_keepDefault_" + timings[0].Name + "_" + timings[0].SuffixAcSpecName;
                }
                else
                {
                    //nWireRow.TestName = NwireInstance.FreeRunClkEnable  + "_" + defalutNwire.Name + "_Port" + "_" + setting;
                    nWireRow.TestName = NwireInstance.FreeRunClkEnable + "_" + timing.Name + "_" + timing.SuffixAcSpecName;
                }

                nWireRow.DcCategory = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.FindNwireCategory(out EnumMessageLevel msgLevel, out string errorMsg);
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    //Report error msg
                    Response.Report(errorMsg, msgLevel, 45);
                }

                nWireRow.DcSelector = "Typ";
                nWireRow.AcCategory = acCategory;
                nWireRow.AcSelector = "Typ";
                nWireRow.PinLevels = HardIpConstData.LevelNwire;
                nWireRow.TimeSets = HardIpConstData.TimesetNwire;
                bool hasUltraFLex = TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);

                if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
                {
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(NwireInstance.ControlFreeRunningClkCs, "");
                    nWireRow.VbtType = function.Type;
                    nWireRow.VbtName = function.FunctionName;
                    nWireRow.ArgList = function.Parameters;
                    function.SetParamValue("enable", "TRUE");
                    function.SetParamValue("pins", hasUltraFLex ? defalutNwire.CreatePortName(EnumEquipment.UltraFlexPlus) : defalutNwire.CreatePortName(EnumEquipment.UltraFlex));
                    nWireRow.Args = function.ArgList;
                }
                else
                {
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(NwireInstance.FreeRunClkEnable, "");
                    nWireRow.VbtType = function.Type;
                    nWireRow.VbtName = function.FunctionName;
                    nWireRow.ArgList = function.Parameters;
                    function.SetParamValue("Portname", defalutNwire.CreatePinNameWithDiff() + "_Port");
                    nWireRow.Args = function.ArgList;
                }
                result.Add(nWireRow);
            }
            return result;
        }

        public string GetAcCategoryByAc(HardIpPattern pattern, string blockName, string timeSets = "")
        {
            List<Timing> timings = pattern.GetTimingsByAc();

            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets, BlockType.HardIp);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets);
            }

            if (timings.Count > 0)
            {
                string category = "";
                if (timeSet2Cat == "" || timeSet2Cat == "TBD")
                {
                    category = blockName + "_" + timeSet2Cat + timings.Aggregate("",
                        (current, timing) => current + timing.Name + "_" + timing.SuffixAcSpecName + "_");
                }
                else
                {
                    category = blockName + "_" + timeSet2Cat + "_" + timings.Aggregate("",
                        (current, timing) => current + timing.Name + "_" + timing.SuffixAcSpecName + "_");
                }

                return category.Trim('_');
            }

            return "";
        }
    }
}
