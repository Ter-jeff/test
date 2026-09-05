using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using Automation.Singleton;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.Settings;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business
{
    [ExcludeFromCodeCoverage]
    public class NwirePortMapLcd : NwirePortMap
    {
        public override List<PortMapSheet> GenerateFlow()
        {
            var portMaps = new List<PortMapSheet>();
            PortMapSheet portMap = new PortMapSheet(SheetName);
            GenPortMapRowNwire(portMap);

            //Add ADG1414 & Get from non Frc sheet
            List<PortSet> rows = AddNonFrc();
            if (rows != null)
            {
                foreach (PortSet row in rows)
                {
                    portMap.AddRow(row);
                }
            }

            PortSet set = new PortSet("");
            set.AddPortRow(new PortRow());
            portMap.AddRow(set);

            AddUart(portMap);
            portMaps.Add(portMap);
            return portMaps;
        }

        private List<PortSet> AddNonFrc()
        {
            List<NonFrcNWires> nonFrcsetting = NwireSingleton.Instance().NonFrcSetting;
            if (nonFrcsetting == null)
            {
                return null;
            }

            List<PortSet> portSetList = new List<PortSet>();
            var list = nonFrcsetting.GroupBy(x => x.PortName).ToList();
            foreach (IGrouping<string, NonFrcNWires> set in list)
            {
                var portSet = new PortSet(set.Key);
                foreach (NonFrcNWires item in set)
                {
                    PortRow row = new PortRow
                    {
                        ProtocolFamily = NWire,
                        ProtocolSettings = NonFrcNWires.CreateSettingName(NwireSingleton.Instance().NodeAdg1414, "name") + "="
                        + NonFrcNWires.CreateSettingName(NwireSingleton.Instance().NodeAdg1414, "defaultValue"),
                        ProtocolSettingValues = new List<string> { NonFrcNWires.CreateSettingdefaultValue(NwireSingleton.Instance().NodeAdg1414) },
                        ProtocolType = item.ProtocalType,
                        FunctionName = item.FunctionName,
                        FunctionPin = item.DeviecPinName
                    };
                    portSet.AddPortRow(row);
                }
                portSetList.Add(portSet);
            }
            return portSetList;
        }
    }
}
