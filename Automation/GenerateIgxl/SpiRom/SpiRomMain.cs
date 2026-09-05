using System;

using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

namespace Automation.GenerateIgxl.SpiRom
{
    public class SpiRomMain : WorkFlowBase<ParaData>
    {
        public override bool PreCheckFlow(ParaData paraData)
        {
            return true;
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                Response.Report("Parsing SPI_ROM Table~", EnumMessageLevel.General, 30);
                GenSpiRom spiRom = new GenSpiRom(EpWorkbook.SpiRomWorkBook);
                spiRom.Workflow();

                Response.Report("Parsing SPI_ROM Table Done!", EnumMessageLevel.General, 50);
            }
            catch (Exception e)
            {
                string message = "SPI_ROM AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }
    }
}
