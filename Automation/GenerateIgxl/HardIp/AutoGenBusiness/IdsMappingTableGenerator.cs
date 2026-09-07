using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using LogLib.Utility;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class IdsMappingTableGenerator
    {
        private readonly IdsMappingSheet _idsMappingFromPlan;

        public IdsMappingTableGenerator(IdsMappingSheet idsMappingFromPlan)
        {
            _idsMappingFromPlan = idsMappingFromPlan;
        }

        public void WorkFlow()
        {
            try
            {
                if (_idsMappingFromPlan == null || !_idsMappingFromPlan.Rows.Any() || _idsMappingFromPlan.IndexSubBlock == -1)
                {
                    return;
                }

                var txt = new List<string>();
                IEnumerable<IGrouping<string, IdsMappingRow>> groups = _idsMappingFromPlan.Rows.GroupBy(x => $"Stage:{x.Stage}, Inst:{x.InstanceName}");

                WriteHeader(txt);
                foreach (IGrouping<string, IdsMappingRow> group in groups)
                {
                    bool first = true;
                    foreach (IdsMappingRow row in group)
                    {
                        WriteContent(row, txt, first);
                        first = false;
                    }
                }

                if (!Directory.Exists(FolderStructure.DirCommon))
                {
                    Directory.CreateDirectory(FolderStructure.DirCommon);
                }

                using (StreamWriter sw = File.CreateText(Path.Combine(FolderStructure.DirCommon, "IDS_Mapping_Table.txt")))
                {
                    foreach (string line in txt)
                    {
                        sw.WriteLine(line);
                    }
                }
                if (!TestProgram.NonIgxlSheetsList.SheetList.Exists(x => x == Path.Combine(FolderStructure.DirCommon, "IDS_Mapping_Table")))
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, "IDS_Mapping_Table");
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show("Generate Ids Mapping Table Error! :");
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private void WriteHeader(List<string> txt)
        {
            var headers = new List<string>
            {
                IdsMappingTableReader.ConHeaderStage,
                "InstanceName",
                IdsMappingTableReader.ConHeaderPinname,
                IdsMappingTableReader.ConHeaderEfusefieldname
            };
            txt.Add(string.Join("\t", headers));
        }

        private void WriteContent(IdsMappingRow idsMappingRow, List<string> txt, bool printFistInfo)
        {
            List<string> content;
            if (printFistInfo)
            {
                string instanceName = idsMappingRow.InstanceName;
                if (string.IsNullOrEmpty(instanceName))
                {
                    instanceName = idsMappingRow.SubBlock;
                }

                content = new List<string>
                {
                    idsMappingRow.Stage,
                    instanceName,
                    idsMappingRow.Pinname,
                    idsMappingRow.Efusefieldname
                };
            }
            else
            {
                content = new List<string>
                {
                    "",
                    "",
                    idsMappingRow.Pinname,
                    idsMappingRow.Efusefieldname
                };
            }
            txt.Add(string.Join("\t", content));
        }
    }
}
