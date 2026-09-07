using System.Collections.Generic;

using CommonLib.Enums;

using IgxlLib;

using LogLib.Static;

namespace Automation.GenerateIgxl.PostAction.GenIgxlProj
{
    internal sealed class NoOpIgxlPackager : IIgxlPackager
    {
        private readonly string _reason;

        public NoOpIgxlPackager(string reason)
        {
            _reason = reason;
        }

        public void GenIgxlProg(List<string> sourceFile, string outputFolder, string projectName, IgxlWorkBook igxlWorkBook, bool isUnitTest)
        {
            Response.Report(
                $"IGXL packaging skipped ({_reason}). Run IGLinkCL.exe manually on the generated .igxlProj.",
                EnumMessageLevel.EndPoint);
        }
    }
}
