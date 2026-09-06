using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.TempMon.Data;
using Automation.GenerateIgxl.PostAction.TempMon.Writer;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.NonIgxlSheets;

namespace Automation.GenerateIgxl.PostAction.TempMon
{
    public class TempMonMain
    {
        private ExcelWorksheet _tempMonConfigSheet;
        private HashSet<TempMonData> _tempMonDatas = new HashSet<TempMonData>();
        private NonIgxlSheets _nonIgxlSheetsList;
        private string _outputFolder = "";

        public TempMonMain(ExcelWorksheet tempMonConfigSheet, HashSet<TempMonData> tempMonDatas, NonIgxlSheets nonIgxlSheetsList, string outputFolder)
        {
            _tempMonConfigSheet = tempMonConfigSheet;
            _tempMonDatas = tempMonDatas;
            _nonIgxlSheetsList = nonIgxlSheetsList;
            _outputFolder = outputFolder;
        }

        public void WorkFlow()
        {
            if (_tempMonDatas == null || !_tempMonDatas.Any())
            {
                return;
            }

            if (_tempMonConfigSheet != null)
            {
                AddTempMonConfigSheet();
            }

            AddTempMonExecutionSheet();
        }

        private void AddTempMonConfigSheet()
        {
            _tempMonConfigSheet.ExportWorkBook2Txt(_outputFolder);
            _nonIgxlSheetsList.Add(_outputFolder, _tempMonConfigSheet.Name);
        }

        private void AddTempMonExecutionSheet()
        {
            TempMonConditionSheet tempMonConditionSheet = new TempMonConditionSheet(_tempMonDatas);
            tempMonConditionSheet.Write(_outputFolder);
            _nonIgxlSheetsList.Add(_outputFolder, TempMonConditionSheet.SheetName);
        }
    }
}
