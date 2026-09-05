using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using System.Text.RegularExpressions;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using LogLib.Base;
using LogLib.Static;

using TestPlanLib.Const;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business
{
    public class NwireInstance
    {
        #region Field
        public const string StartSupportBoard = "StartSBClock";
        public const string StopSupportBoard = "StopSBClock";
        public const string ControlSupportBoardCs = "ControlSupportBoardClock";
        public const string FreeRunClkEnable = "FreeRunClk_Enable";
        public const string FreeRunClkDisable = "FreeRunClk_Disable";
        public const string ControlFreeRunningClkCs = "ControlFreeRunningClock";
        public const string RelayControl = "Relay_Control";
        public const string ControlRelayCs = "ControlRelay";
        public const string ControlDefaultDisable = "Default_Disable";

        protected const string Typ = "Typ";

        protected string NWireDcCategory;
        protected string NWireAcCategory;
        protected string NWireLevel;
        protected string NWireTimeSet;
        protected IProgress<Progress> Progress;

        #endregion

        #region Constructor

        public NwireInstance()
        {
            NWireDcCategory = GetNwireDcCategoryName();
            NWireAcCategory = "Common";
            NWireTimeSet = "TIMESET_nWire";
            NWireLevel = "Levels_nWire";
        }

        #endregion

        public virtual List<InstanceRow> GenerateFlow()
        {
            List<InstanceRow> rows = new List<InstanceRow>();
            var flexJobs = new List<string>();
            if (TestPlanStatic.JobInfoSheet != null)
            {
                flexJobs.AddRange(TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.ToUpper().Equals("UF")).Select(x => x.JobName));
            }
            else if (TestPlanStatic.Equipments.First().Equals(EnumEquipment.UltraFlex))
            {
                flexJobs.Add("All");
            }

            if (flexJobs.Any())
            {
                //Start Support board 
                if (TestPlanStatic.UfInstanceTable?.CheckExist(StartSupportBoard) != true)
                {
                    InstanceRow startSbcInstance = CreateStartSbcInstance();
                    rows.Add(startSbcInstance);
                }
                InstanceRow stopSbcInstance = CreateStopSbcInstance();
                rows.Add(stopSbcInstance);
            }

            List<string> allRelays = new List<string>();
            List<string> onRelays = new List<string>();
            List<string> offRelays = new List<string>();

            DataTable table = NwireSingleton.Instance().SettingInfo.SettingTable;
            List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;
            foreach (EnumEquipment testerType in TestPlanStatic.Equipments)
            {
                int pinCount = 1;
                foreach (ProtocolAwarePin pin in nWirePin)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        string item = table.Rows[i][0].ToString();
                        pin.FlowControlAction = item;
                        pin.ControlAction = table.Rows[i][pinCount].ToString();

                        if (item.Contains("Disable"))
                        {
                            //Create Disable instance
                            InstanceRow disable = CreateDisableInstance(pin, testerType);
                            rows.Add(disable);
                        }
                        else if (item.Equals("Enable"))
                        {
                            //Create Enable instance
                            InstanceRow enable = CreateEnableInstance(pin, testerType);
                            rows.Add(enable);
                        }
                        else
                        {
                            InstanceRow disable = CreateDisableInstance(pin, testerType);
                            rows.Add(disable);
                            InstanceRow enable = CreateEnableInstance(pin, testerType);
                            rows.Add(enable);
                        }
                    }

                    allRelays.Add(pin.RelayControl);
                    if (pin.Freq < ProtocolAwarePin.MinFreq)
                    {
                        onRelays.Add(pin.RelayControl);
                    }
                    else
                    {
                        offRelays.Add(pin.RelayControl);
                    }

                    pinCount++;

                }
            }
            return rows;
        }

        /// <summary>
        /// Start Support Board Clock
        /// </summary>
        /// <returns></returns>
        protected InstanceRow CreateStartSbcInstance()
        {
            InstanceRow row = new InstanceRow { TestName = StartSupportBoard };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ControlSupportBoardCs, "basic", true);
            if (function.IsFound)
            {
                function.SetParamValue("enable", "TRUE");
                function.SetParamValue("frequency", NwireSingleton.Instance().SettingInfo.SupportBoardFreq.ToString());
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(StartSupportBoard, "basic");
                function.ArgList[0] = NwireSingleton.Instance().SettingInfo.SupportBoardFreq.ToString();
            }
            row.VbtType = function.Type;
            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;

            return row;
        }

        /// <summary>
        /// Stop Support Board Clock
        /// </summary>
        /// <returns></returns>
        protected InstanceRow CreateStopSbcInstance()
        {
            InstanceRow row = new InstanceRow { TestName = StopSupportBoard };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ControlSupportBoardCs, "basic", true);
            if (function.IsFound)
            {
                function.SetParamValue("enable", "FALSE");
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(StopSupportBoard, "basic");
            }
            row.VbtType = function.Type;
            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;

            return row;
        }

        /// <summary>
        /// Enable nWire pin
        /// </summary>
        /// <param name="awarePin"></param>
        /// <returns></returns>
        protected InstanceRow CreateEnableInstance(ProtocolAwarePin awarePin, EnumEquipment testerType)
        {
            InstanceRow row = new InstanceRow
            {
                TestName = FreeRunClkEnable
                + "_" + awarePin.CreatePinNameWithDiff()
                + "_" + (testerType.Equals(EnumEquipment.UltraFlex) ? "UF" : "UFP")
                + (string.IsNullOrEmpty(awarePin.FlowControlAction) ? "" : "_" + awarePin.FlowControlAction)
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ControlFreeRunningClkCs, "basic", true);
            if (function.IsFound)
            {
                function.SetParamValue("enable", "TRUE");
                function.SetParamValue("pins", awarePin.CreatePortName(testerType));

                string configFreq;
                configFreq = ResolveConfigFreq(awarePin);
                function.SetParamValue("frequency", configFreq);
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(FreeRunClkEnable, "basic");
                function.SetParamValue("PortName", awarePin.CreatePortName(EnumEquipment.UltraFlex));
            }
            row.ColumnA = awarePin.OutClk;
            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            row.VbtType = function.Type;
            row.DcCategory = !string.IsNullOrEmpty(awarePin.DcCategory) ? awarePin.DcCategory : NWireDcCategory;
            row.DcSelector = Typ;
            if (testerType == EnumEquipment.UltraFlex)
            {
                row.AcCategory = NWireAcCategory + (string.IsNullOrEmpty(awarePin.FlowControlAction) ? "" : "_" + awarePin.FlowControlAction);
                row.AcSelector = Typ;
                row.TimeSets = NWireTimeSet;
            }
            row.PinLevels = NWireLevel;
            row.Args = function.ArgList;

            return row;
        }
        internal static string ResolveConfigFreq(ProtocolAwarePin awarePin)
        {
            if (Regex.IsMatch(awarePin.ControlAction, "Enable@", RegexOptions.IgnoreCase))
            {
                string frequency = awarePin.ControlAction;
                string valueFreq = Regex.Match(frequency, TestPlanConst.UnitRegPattern)
                                        .Groups[TestPlanConst.Value].ToString();
                string unit = Regex.Match(frequency, TestPlanConst.UnitRegPattern)
                                   .Groups[TestPlanConst.Unit].ToString();

                if (valueFreq.TryCombineHz(unit, out string targetFreq))
                {
                    return targetFreq;
                }
            }

            return awarePin.Freq.ToString();
        }
        /// <summary>
        /// Disable nWire Pin
        /// </summary>
        /// <param name="awarePin"></param>
        /// <returns></returns>
        protected InstanceRow CreateDisableInstance(ProtocolAwarePin awarePin, EnumEquipment testerType)
        {
            InstanceRow row = new InstanceRow
            {
                TestName = FreeRunClkDisable
                + "_" + awarePin.CreatePinNameWithDiff()
                + "_" + (testerType.Equals(EnumEquipment.UltraFlex) ? "UF" : "UFP")
                + (string.IsNullOrEmpty(awarePin.FlowControlAction) ? "" : "_" + awarePin.FlowControlAction)
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ControlFreeRunningClkCs, "basic", true);
            if (function.IsFound)
            {
                function.SetParamValue("enable", "FALSE");
                function.SetParamValue("pins", awarePin.CreatePortName(testerType));
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(FreeRunClkDisable, "");
                function.SetParamValue("PortName", awarePin.CreatePortName(EnumEquipment.UltraFlex));
            }
            row.ColumnA = awarePin.OutClk;
            row.VbtType = function.Type;
            row.VbtName = function.FullFunctionName;
            row.DcCategory = !string.IsNullOrEmpty(awarePin.DcCategory) ? awarePin.DcCategory : NWireDcCategory;
            row.DcSelector = Typ;
            if (testerType == EnumEquipment.UltraFlex)
            {
                row.AcCategory = NWireAcCategory + (string.IsNullOrEmpty(awarePin.FlowControlAction) ? "" : "_" + awarePin.FlowControlAction);
                row.AcSelector = Typ;
                row.TimeSets = NWireTimeSet;
            }
            row.PinLevels = NWireLevel;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;

            return row;
        }

        private string GetNwireDcCategoryName()
        {
            string category = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.FindNwireCategory(out EnumMessageLevel msgLevel, out string errorMsg);
            if (!string.IsNullOrEmpty(errorMsg))
            {
                //Report error msg
                Response.Report(errorMsg, msgLevel, 90);
            }
            return category;

        }

    }
}
