using System.Collections.Generic;

using Automation.GenerateIgxl.PreAction.GenPinMap;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class IoPinMapReader
    {
        public static void ReadPinMap()
        {
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap == null)
            {
                if (EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinMap"] != null && EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinGroup"] != null && EpWorkbook.TestPlanWorkbook.Worksheets["IO_Continuity"] != null)
                {
                    var readPinMapFlow = new PinMapMain();
                    PinMapSheet ioPinMap = new ReadPinMapSheet().ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinMap"]);
                    ioPinMap.Name = "PinMap";
                    ExcelWorksheet ioPinGroup = EpWorkbook.TestPlanWorkbook.Worksheets["IO_PinGroup"];
                    IoContiSheet ioContinuity = new IoContiReader().ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["IO_Continuity"]);
                    PinMapSheet sheet = readPinMapFlow.WorkFlow(ioPinMap, ioPinGroup, ioContinuity);
                    if (sheet != null)
                    {
                        TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, sheet);
                    }
                    else
                    {
                        sheet = new PinMapSheet("PinMap");
                        TestProgram.IgxlWorkBk.PinMapPair = new KeyValuePair<string, PinMapSheet>(FolderStructure.DirPinMap, sheet);
                    }
                }
                else
                {
                    ErrorReportManager.AddError(PreActionErrorType.E_MissingDocument_01, "", 1, 0, "Missing IO_PinMap/IO_Continuity sheet in TestPlan");
                }
            }
        }
    }
}
