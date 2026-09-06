using System.Collections.Generic;
using System.IO;

using Automation.Static;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using LocalSpecs = Cautogen.AutoCZ.CharPostProcessor.LocalSpec.LocalSpecs;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class DfcSheetGenerator
    {


        /* Member function */
        public static void Generate()
        {
            GeneralFunc.WriteMessage("Generating DFC sheet... ");
            if (File.Exists(Path.Combine(LocalSpecs.OutputFolder, "IGLink", "trunk", "Common", "Common_Sheets", "DFC_List.txt")))
            {
                File.Delete(Path.Combine(LocalSpecs.OutputFolder, "IGLink", "trunk", "Common", "Common_Sheets", "DFC_List.txt"));
            }
            string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);


            // create blank instance sheet for each char plan sheet
            //var czInstSheet = new InstanceSheet("TestInst_CZ_" + planSheet.SheetName);

            // export cz inst sheet
            string czFileName = Path.Combine(outputFolder, "DFC_List.txt");
            WriteDfc(czFileName);
        }

        public static void WriteDfc(string fileName)
        {
            var lines = new List<string> { "Test Instance" };
            for (int i = 2; i <= LocalSpecs.DFCSheet.Dimension.Rows; i++)
            {
                string instance = LocalSpecs.DFCSheet.Cells[i, 1].Text;
                lines.Add(instance);
            }
            File.WriteAllLines(fileName, lines);
            TestProgram.NonIgxlSheetsList.Add(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName));
            LocalSpecs.GenOthers.Add(fileName);
        }
    }
}
