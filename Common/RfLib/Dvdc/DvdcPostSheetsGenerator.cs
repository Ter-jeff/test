using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.InputManager.Data;
using Automation.Static;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using RfLib.Dvdc.GenVbt;

namespace RfLib.Dvdc
{
    internal static class DvdcPostSheetsGenerator
    {
        internal static SubFlowSheet Generate(HardIpInputData hardIpInputData, List<InstanceSheet> instanceSheets, BinTableSheet binTableSheet)
        {
            Response.Report("Generating Relay Instance ...", percentage: 75);
            new RelayInstanceGenerator().GenRelayInstance(hardIpInputData.PlanDic);

            Response.Report("Generating Nwire Instance ...", percentage: 80);
            new NwireInstanceGenerator().GenNwireInstance(hardIpInputData.PlanDic, instanceSheets);

            Response.Report("Generating AcCategory ...", percentage: 82);
            new AcCategoryGenerator().GenAcCategory(hardIpInputData.PlanDic, instanceSheets);

            Response.Report("Generating Init Flag Flow ...", percentage: 90);
            SubFlowSheet initFlow = new InitFlagGenerator().GenInitFlag(hardIpInputData.PlanDic, binTableSheet);

            Response.Report("Generating Related information in VBT Files ...", percentage: 92);

            Response.Report("Adding Sheets into Project Object ...", percentage: 95);
            WirelessVBTGenerator.WorkFlow(hardIpInputData.PlanDic, LocalSpecs.TestPlanFileName);

            if (hardIpInputData.HardIpRegAssigns.Count > 0)
            {
                new RegAssignGenerator(hardIpInputData).WorkFlow(hardIpInputData.HardIpRegAssigns);
            }

            //if (hardIpInputData.InterposeAssigns.Count > 0)
            //{
            //    new InterposeAssignGenerator().WorkFlow(hardIpInputData.InterposeAssigns);
            //}

            return initFlow;
        }
    }
}
