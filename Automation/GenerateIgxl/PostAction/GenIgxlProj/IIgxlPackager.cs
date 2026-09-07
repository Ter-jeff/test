using System.Collections.Generic;

using IgxlLib;

namespace Automation.GenerateIgxl.PostAction.GenIgxlProj
{
    /// <summary>
    /// Builds IGLink projects (.igxlProj) for the current Autogen run.
    /// Implementations may run in-process or shell out to a side-car (e.g. Automation.IgxlPackaging.exe).
    /// </summary>
    public interface IIgxlPackager
    {
        void GenIgxlProg(List<string> sourceFile, string outputFolder, string projectName, IgxlWorkBook igxlWorkBook, bool isUnitTest);
    }
}
