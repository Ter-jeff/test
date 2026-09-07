using System.IO;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class GlobalSpecGenerator
    {
        public static void Generate()
        {
            var globalSpecSheet = new GlobalSpecSheet("Global Specs", false);
            foreach (GlobalSpec row in LocalSpecs.TestProgram.GlobalSpecRows)
            {
                globalSpecSheet.AddRow(row);
            }

            string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.GlobalFolder);

            string globalSpecFile = Path.Combine(outputFolder, "Global Specs.txt");

            globalSpecSheet.Write(globalSpecFile, "2.0");
            LocalSpecs.GenSheets.Add(globalSpecSheet);
        }
    }
}
