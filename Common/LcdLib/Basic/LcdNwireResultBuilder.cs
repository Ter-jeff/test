using Automation.GenerateIgxl.Basic.Business.GenNwire.Base;
using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;
using Automation.Singleton;
using Automation.Static;

namespace LcdLib.Basic
{
    internal static class LcdNwireResultBuilder
    {
        internal static NwireResult? Generate()
        {
            if (!NwireSingleton.Instance().HasNwirePin)
            {
                return null;
            }

            var nwireResult = new NwireResult();

            NwireFlow flow = new NwireFlow();
            nwireResult.NWireFlowSheets = flow.GenerateFlow();

            NwireInstance instance = new NwireInstance();
            nwireResult.NWireInstanceRows = instance.GenerateFlow();

            if (TestProgram.IgxlWorkBk.ChannelMapSheets != null && TestProgram.IgxlWorkBk.ChannelMapSheets.Count != 0)
            {
                NwireSingleton.Instance().SetSuperClock(TestProgram.IgxlWorkBk.ChannelMapSheets);
            }

            NwirePortMapLcd portMap = new NwirePortMapLcd();
            nwireResult.PortMapSheets = portMap.GenerateFlow();

            NwireBintable bintable = new NwireBintable();
            nwireResult.BinTables = bintable.GenerateFlow();

            return nwireResult;
        }
    }
}
