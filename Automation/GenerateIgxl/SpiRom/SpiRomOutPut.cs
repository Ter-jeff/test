using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.SpiRom
{
    public class SpiRomOutPut
    {
        public string DcSpecsSheetName = "";

        public List<SubFlowSheet> FlowList = new List<SubFlowSheet>();
        public List<InstanceSheet> InstanceList = new List<InstanceSheet>();
        public List<TimeSetBasicSheet> TimeSetList = new List<TimeSetBasicSheet>();
        public List<LevelSheet> LevelList = new List<LevelSheet>();
        public List<PatSetSheet> PatSetlList = new List<PatSetSheet>();
        public List<BinTableRow> BinTableRows = new List<BinTableRow>();
        public string SpiromCodeFileString;

        private string _outPutPath = "";

        public void AddDataToLocalSpace()
        {
            _outPutPath = FolderStructure.DirSpiRom;

            AddFlow();

            AddInstance();

            AddTimeSet();

            AddLevel();

            AddPatSet();

            AddBinTableRows();
        }

        private void AddFlow()
        {
            foreach (SubFlowSheet flow in FlowList)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(_outPutPath, flow);
            }
        }

        private void AddInstance()
        {
            foreach (InstanceSheet instance in InstanceList)
            {
                TestProgram.IgxlWorkBk.AddInsSheet(_outPutPath, instance);
            }
        }

        private void AddBinTableRows()
        {
            BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            foreach (BinTableRow binTableRow in BinTableRows)
            {
                binTable.AddRow(binTableRow);
            }
        }

        private void AddTimeSet()
        {
            foreach (TimeSetBasicSheet timeSetSheet in TimeSetList)
            {
                if (TestProgram.IgxlWorkBk.TimeSetSheets.Values.ToList().Exists(x => x.Name.Equals(timeSetSheet.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    TestProgram.IgxlWorkBk.TimeSetSheets.Remove(TestProgram.IgxlWorkBk.TimeSetSheets.FirstOrDefault(x => x.Value.Name.Equals(timeSetSheet.Name, StringComparison.OrdinalIgnoreCase)).Key);
                }
                TestProgram.IgxlWorkBk.AddTimeSetSheet(_outPutPath, timeSetSheet);
            }
        }

        private void AddLevel()
        {
            foreach (LevelSheet levelSheet in LevelList)
            {
                if (TestProgram.IgxlWorkBk.LevelSheets.Values.ToList().Exists(x => x.Name.Equals(levelSheet.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    TestProgram.IgxlWorkBk.LevelSheets.Remove(TestProgram.IgxlWorkBk.LevelSheets.FirstOrDefault(x => x.Value.Name.Equals(levelSheet.Name, StringComparison.OrdinalIgnoreCase)).Key);
                }
                TestProgram.IgxlWorkBk.AddLevelSheet(_outPutPath, levelSheet);
            }
        }

        private void AddPatSet()
        {
            foreach (PatSetSheet patSetSheet in PatSetlList)
            {
                TestProgram.IgxlWorkBk.AddPatSetSheet(_outPutPath, patSetSheet);
            }
        }

        public void OutPutNonIgxlToTxt(string pStrPath, string pStrSheetName)
        {
            string lStrFileName = Path.Combine(pStrPath, pStrSheetName + ".txt");

            CreateFilePath(pStrPath);

            if (File.Exists(lStrFileName))
            {
                File.Delete(lStrFileName);
            }

            var recordData = new StreamWriter(lStrFileName);

            recordData.Write(SpiromCodeFileString);

            TestProgram.NonIgxlSheetsList.Add(pStrPath, pStrSheetName);

            recordData.Close();

        }

        private void CreateFilePath(string pFilePath)
        {
            if (!Directory.Exists(pFilePath))
            {
                Directory.CreateDirectory(pFilePath);
            }
        }
    }
}
