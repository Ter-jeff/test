using System.Collections.Generic;
using System.IO;

using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace Automation.GenerateIgxl.PostAction.SelSram
{
    public class SelSramMain
    {
        public void WorkFlow()
        {
            if ((BlockStatus.GetAutomationBlockStatus(BlockStatus.Scan).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Scan)) ||
                   (BlockStatus.GetAutomationBlockStatus(BlockStatus.Mbist).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Mbist)))
            {
                var selSram = SelSramPatternSingleton.GetInstance();
                selSram.ReadbackSheet();

                if (selSram.DicReadbackPat.Count != 0)
                {
                    selSram.FillData();
                    const string outputFile = "SelSram_Chklist.txt";
                    selSram.Write2Txt(Path.Combine(FolderStructure.DirCommonSheets, outputFile));
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, Path.GetFileNameWithoutExtension(outputFile));

                    Response.Report("Start to generate SelSram sheets");
                    var enableGenList = new List<string>();
                    var ssWriter = SelSramWriterSingleton.GetInstance();
                    var selSramInstanceSheet = new InstanceSheet("TestInst_SelSram");
                    ssWriter.WriteInstanceSheet(ref selSramInstanceSheet, ref enableGenList);

                    var selSramFlowSheet = new SubFlowSheet("Flow_SelSram", "SLESRM");
                    ssWriter.WriteFlowSheet(ref selSramFlowSheet, enableGenList);

                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirCommonSheets, selSramFlowSheet);
                    TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, selSramInstanceSheet);
                    Response.Report("Generate SelSram sheets completed !");
                }
            }
        }
    }
}
