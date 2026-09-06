using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business
{
    public class NwirePortMap
    {
        protected const string SheetName = "PortMap";
        protected const string ClockOut = "ClockOut";
        protected const string ClockOut2 = "ClockOut_2";
        protected const string Refclk = "Refclk";
        protected const string Clock = "Clock";
        protected const string ClockDiff = "Clock_Diff";
        protected const string NWire = "nWire";
        protected const string IgnoreContextChanges = "IgnoreContextChanges=TRUE";

        public virtual List<PortMapSheet> GenerateFlow()
        {
            var portMaps = new List<PortMapSheet>();
            foreach (EnumEquipment equipment in TestPlanStatic.Equipments)
            {
                string testType = equipment == EnumEquipment.UltraFlex ? "UF" : "UFP";
                string sheetName = SheetName + "_" + testType;
                var portMap = new PortMapSheet(sheetName);

                if (equipment.Equals(EnumEquipment.UltraFlex))
                {
                    GenPortMapRowNwire(portMap);
                }

                AddUart(portMap);
                portMaps.Add(portMap);
            }
            return portMaps;
        }

        protected void GenPortMapRowNwire(PortMapSheet portMap)
        {
            PortSet portSet;
            PortRow clockRow;
            PortRow refclkRow;
            List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;

            foreach (ProtocolAwarePin awarePin in nWirePin)
            {
                string type = SettingPinType(awarePin);
                portSet = new PortSet(awarePin.CreatePortName(EnumEquipment.UltraFlex));
                (string Family, string Type) proto = ParseProtocolOrDefault(awarePin.Protocol, NWire, type);
                clockRow = new PortRow
                {
                    FunctionName = ClockOut,
                    FunctionPin = awarePin.CreatePaClkPinName(EnumEquipment.UltraFlex),
                    ProtocolFamily = proto.Family,
                    ProtocolType = proto.Type,
                    ProtocolSettings = IgnoreContextChanges,
                };
                clockRow.AddSetting("True");
                portSet.AddPortRow(clockRow);

                #region extraPin

                if (!string.IsNullOrEmpty(awarePin.ExtraPin))
                {
                    //Add SPMI_SCLK(ClockOut_SPMI)
                    const string regPin = @"^(?<pin>[\w]+)[\s]*([\(](?<name>[^)]+)[\)])?";
                    foreach (string item in awarePin.ExtraPin.Split(','))
                    {
                        string pin = Regex.Match(item, regPin, RegexOptions.IgnoreCase).Groups["pin"].ToString().Trim();
                        string name = Regex.Match(item, regPin, RegexOptions.IgnoreCase).Groups["name"].ToString().Trim();
                        var extraPin = new PortRow();
                        extraPin = new PortRow
                        {
                            FunctionName = name,
                            FunctionPin = pin,
                            ProtocolFamily = proto.Family,
                            ProtocolType = proto.Type,
                            ProtocolSettings = IgnoreContextChanges
                        };
                        extraPin.AddSetting("True");
                        portSet.AddPortRow(extraPin);
                    }
                }

                #endregion

                refclkRow = new PortRow
                {
                    FunctionName = Refclk,
                    FunctionPin = awarePin.RefClk,
                    ProtocolFamily = proto.Family,
                    ProtocolType = proto.Type,
                    ProtocolSettings = IgnoreContextChanges
                };
                refclkRow.AddSetting("True");
                portSet.AddPortRow(refclkRow);

                if (awarePin.PinType == IoPinType.Diff)
                {
                    //Add XO0_PA
                    var outclkDiff = new PortRow
                    {
                        FunctionName = ClockOut2,
                        FunctionPin = awarePin.CreatePaClkDiffPinName(EnumEquipment.UltraFlex),
                        //outclkDiff.ProtocolFamily = NWire;
                        ProtocolFamily = proto.Family,
                        ProtocolType = proto.Type,
                        ProtocolSettings = IgnoreContextChanges
                    };
                    outclkDiff.AddSetting("True");
                    portSet.AddPortRow(outclkDiff);
                }

                portMap.AddRow(portSet);
            }
        }

        public static (string Family, string Type) ParseProtocolOrDefault(
            string protocol,
            string defaultA,
            string defaultB)
        {
            if (string.IsNullOrWhiteSpace(protocol))
            {
                return (defaultA, defaultB);
            }

            string[] parts = protocol.Split(':');
            if (parts.Length != 2)
            {
                return (defaultA, defaultB);
            }

            string family = parts[0].Trim();
            string type = parts[1].Trim();
            if (string.IsNullOrWhiteSpace(family) || string.IsNullOrWhiteSpace(type))
            {
                return (defaultA, defaultB);
            }

            return (family, type);
        }

        private string SettingPinType(ProtocolAwarePin awarePin)
        {
            switch (awarePin.PinType)
            {
                case IoPinType.Single:
                    break;
                case IoPinType.Diff:
                    return ClockDiff;
            }
            return Clock;
        }

        protected void AddUart(PortMapSheet portMap)
        {
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            Dictionary<string, string> uartPinDic = pinMap.GetUartPinDic();
            var portSetTx = new PortSet("UART_TX");
            var rowTx = new PortRow { ProtocolFamily = NWire, ProtocolType = "UART_PA_TX", FunctionName = "UART_TX" };
            if (uartPinDic.TryGetValue("UART_TX", out string txPinName))
            {
                rowTx.FunctionPin = txPinName;
            }
            portSetTx.AddPortRow(rowTx);
            if (pinMap.IsGroupExist(portSetTx.PortName))
            {
                portMap.AddRow(portSetTx);
            }

            var portSetRx = new PortSet("UART_RX");

            var rowRx = new PortRow { ProtocolFamily = NWire, ProtocolType = "UART_PA_RX", FunctionName = "UART_RX" };
            if (uartPinDic.TryGetValue("UART_RX", out string rxPinName))
            {
                rowRx.FunctionPin = rxPinName;
            }
            portSetRx.AddPortRow(rowRx);
            if (pinMap.IsGroupExist(portSetRx.PortName))
            {
                portMap.AddRow(portSetRx);
            }
        }
    }
}
