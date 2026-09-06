using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using EfuseCheckCmdLib.EFuse.EFuseApp;

namespace EfuseCheckCmdLib.CFGTable
{
    public partial class CfgTableReader
    {

        [GeneratedRegex(".txt", RegexOptions.IgnoreCase)]
        private static partial Regex TxtExtensionRegex();
        public static readonly Dictionary<string, EfuseCfgTable> CfgTable = [];
        public static bool IsContainCfgTable { get; private set; }

        public static void WorkFlow(string path)
        {
            if (TxtExtensionRegex().IsMatch(path))
            {
                try
                {
                    var table = new EfuseCfgTable();
                    table.ReadTxt(path);
                    CfgTable.Add(Path.GetFileNameWithoutExtension(path), table);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format(ex.ToString()));
                }
            }
            IsContainCfgTable = CfgTable.Count != 0;
        }

        public static void Reset()
        {
            CfgTable.Clear();
        }

        public static EfuseCfgTable CfgTableSel()
        {
            return CfgTable.ElementAt(0).Value;
        }
    }
}
