using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.Static;

using LogLib.Utility;

using OfficeOpenXml;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public partial class EfuseCfgTableReader1
    {

        [GeneratedRegex(".xlsx", RegexOptions.IgnoreCase)]
        private static partial Regex XlsxExtensionRegex();

        [GeneratedRegex(@"(CFG)+_*[a-zA-Z]*(Cond)+[a-zA-Z]*\w*(Table)+|Config_table", RegexOptions.IgnoreCase)]
        private static partial Regex CfgConditionTableNameRegex();

        [GeneratedRegex(".txt", RegexOptions.IgnoreCase)]
        private static partial Regex TxtExtensionRegex();
        public static readonly Dictionary<string, EfuseCfgTable> CfgTable = [];
        public static bool IsContainCfgTable { get; private set; }

        public static void WorkFlow(string path)
        {
            Reset();
            if (XlsxExtensionRegex().IsMatch(path))
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                using (var ep = new ExcelPackage(fs))
                {
                    EpWorkbook.TestPlanWorkbook = ep.Workbook;
                    foreach (ExcelWorksheet ws in ep.Workbook.Worksheets)
                    {
                        if (CfgConditionTableNameRegex().IsMatch(ws.Name))
                        {
                            var table = new EfuseCfgTable();
                            table.Read(ws);
                            CfgTable.Add(ws.Name, table);
                        }
                    }
                    if (CfgTable.Count == 0)
                    {
                        var cfgGenerator = new GenerateConfigTable(EpWorkbook.TestPlanWorkbook);
                        GenerateConfigTable.WorkFlow();
                        if (cfgGenerator.ConfigTables.Count > 0)
                        {
                            foreach (ExcelWorksheet worksheet in cfgGenerator.ConfigTables)
                            {
                                var table = new EfuseCfgTable();
                                table.Read(worksheet);
                                CfgTable.Add(worksheet.Name, table);
                            }
                        }
                    }

                    ep.Dispose();
                }
                fs.Dispose();
            }
            else if (TxtExtensionRegex().IsMatch(path))
            {
                try
                {
                    var table = new EfuseCfgTable();
                    table.ReadTxt(path);
                    CfgTable.Add(Path.GetFileNameWithoutExtension(path), table);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }
            IsContainCfgTable = CfgTable.Count != 0;
        }

        private static void Reset()
        {
            CfgTable.Clear();
        }
    }
}
