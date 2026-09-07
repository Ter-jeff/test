using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class InterposeAssignGenerator
    {
        public void WorkFlow(List<InterposeAssign> interposeAssigns, string postFixed = "")
        {
            //set up header
            var dt = new DataTable();
            dt.Columns.Add(InterposeAssign.ColSetupName, typeof(string));
            dt.Columns.Add(InterposeAssign.ColInterposes, typeof(string));

            DataRow newRow = dt.NewRow();
            newRow[InterposeAssign.ColSetupName] = InterposeAssign.ColSetupName;
            newRow[InterposeAssign.ColInterposes] = InterposeAssign.ColInterposes;
            dt.Rows.Add(newRow);

            var groupsList = interposeAssigns.GroupBy(x => x.BlockName).ToDictionary(p => p.Key, p => p.OrderBy(x => (int)x.Type).ToList());
            foreach (KeyValuePair<string, List<InterposeAssign>> group in groupsList)
            {
                foreach (InterposeAssign assigns in group.Value)

                {
                    bool isFirst = true;
                    foreach (string assign in assigns.InterposeAssignList)
                    {
                        try
                        {
                            newRow = dt.NewRow();
                            if (isFirst)
                            {
                                newRow[InterposeAssign.ColSetupName] = IsSupportedType(assigns.Type) ? assigns.AssignName : "";
                                newRow[InterposeAssign.ColInterposes] = string.IsNullOrEmpty(assign) ? "NA" : assign;
                                isFirst = false;
                            }
                            else
                            {
                                newRow[InterposeAssign.ColInterposes] = string.IsNullOrEmpty(assign) ? "NA" : assign;

                            }
                            dt.Rows.Add(newRow);
                        }
                        catch (Exception ex)
                        {
                            ErrorMessageBox.Show(string.Format(ex.ToString()));
                        }
                    }
                }
            }

            if (!Directory.Exists(FolderStructure.DirHardIp))
            {
                Directory.CreateDirectory(FolderStructure.DirHardIp);
            }
            using (StreamWriter sw = File.CreateText(Path.Combine(FolderStructure.DirHardIp, "ExternalInterposeAssign.txt")))
            {
                foreach (DataRow row in dt.Rows)
                {
                    sw.WriteLine(string.Join("\t", row.ItemArray));
                }
            }
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirHardIp, "ExternalInterposeAssign");
        }

        private bool IsSupportedType(InterposeAssignType type)
        {
            if (type.Equals(InterposeAssignType.InterposePreInit) || type.Equals(InterposeAssignType.InterposePostInit) || type.Equals(InterposeAssignType.InterposePreRst) || type.Equals(InterposeAssignType.InterposePostRst))
            {
                return false;
            }
            return true;
        }
    }
}
