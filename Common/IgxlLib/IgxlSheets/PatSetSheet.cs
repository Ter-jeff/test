using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class PatSetSheet : IgxlSheet<PatSet>
    {
        public override string SheetType => "DTPatternSetSheet";
        public override string IgxlSheetName => IgxlSheetNames.PatternSet;

        public HashSet<string> ExistPatSet
        {
            get
            {
                return new HashSet<string>(PatSetRowsDic.Keys, StringExtensions.IgnoreCase);
            }
        }

        private Dictionary<string, PatSet> PatSetRowsDic
        {
            get
            {
                return Rows.ToDictionary(x => x.PatSetName, x => x, StringExtensions.IgnoreCase);
            }
        }

        public Dictionary<string, PatSet> PatSetRowDic
        {
            get
            {
                return PatSetRowsDic;
            }
        }

        public PatSetSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public PatSetSheet(string sheetName) : base(sheetName)
        {
        }

        public bool IsExistTheSamePatSet(PatSet patSet, out string patSetName)
        {
            patSetName = string.Empty;
            var candidateKeys = PatSetRowsDic.Keys.Where(key => key != null && key.Contains(patSet.PatSetName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (candidateKeys.Count == 0)
            {
                return false;
            }

            foreach (string key in candidateKeys)
            {
                PatSet row = PatSetRowsDic[key];
                if (row.PatSetRows.Count != patSet.PatSetRows.Count)
                {
                    continue;
                }

                bool isMatch = true;
                for (int i = 0; i < patSet.PatSetRows.Count; i++)
                {
                    if (!patSet.PatSetRows[i].CompareRow(row.PatSetRows[i]))
                    {
                        isMatch = false;
                        break;
                    }
                }

                if (isMatch)
                {
                    patSetName = key;
                    return true;
                }
            }

            return false;
        }

        public bool IsExist(string patSetName)
        {
            return ExistPatSet.Contains(patSetName.ToUpper());
        }

        public PatSetSheet Remove(List<string> patSetNames)
        {
            var patSets = new List<PatSet>();
            foreach (PatSet patset in Rows)
            {
                if (!patSetNames.ToList().Contains(patset.PatSetName, StringExtensions.IgnoreCase))
                {
                    patSets.Add(patset);
                }
            }

            Rows = patSets;
            return this;
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P2(fileName, version);
        }

        protected override void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = GetMaxCount(sheetInfo);
            int patternSetIndex = GetIndexFrom(sheetInfo, "Pattern Set");
            int tDGroupSetIndex = GetIndexFrom(sheetInfo, "TD Group");
            int timeDomainIndex = GetIndexFrom(sheetInfo, "Time Domain");
            int enableIndex = GetIndexFrom(sheetInfo, "Enable");
            int fileGroupNameIndex = GetIndexFrom(sheetInfo, "File/Group Name");
            int burstIndex = GetIndexFrom(sheetInfo, "Burst");
            int startLabelIndex = GetIndexFrom(sheetInfo, "Start Label");
            int stopLabelIndex = GetIndexFrom(sheetInfo, "Stop Label");
            int commentIndex = GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                WriteHeader(arr, sb);
            }
            #endregion

            #region data
            var mainList = Rows.Where(x => x.PatSetRows.All(y => !y.IsBackup)).ToList();
            var backupList = Rows.Where(x => x.PatSetRows.Any(y => y.IsBackup)).ToList();
            if (backupList.Count != 0)
            {
                PrintBackupEmptyRows(mainList);
                mainList.AddRange(backupList);
            }

            for (int index = 0; index < mainList.Count; index++)
            {
                PatSet patSet = mainList[index];
                if (string.IsNullOrEmpty(patSet.PatSetName) && patSet.PatSetRows.Count == 0)
                {
                    sb.AppendLine("");
                    continue;
                }
                for (int i = 0; i < patSet.PatSetRows.Count; i++)
                {
                    PatSetRow row = patSet.PatSetRows[i];
                    string[] arr = [.. Enumerable.Repeat("", maxCount)];
                    arr[0] = row.ColumnA ?? "";
                    arr[patternSetIndex] = patSet.PatSetName;
                    if (tDGroupSetIndex != -1)
                    {
                        arr[tDGroupSetIndex] = row.TdGroup;
                    }

                    arr[timeDomainIndex] = row.TimeDomain;
                    arr[enableIndex] = row.Enable;
                    arr[fileGroupNameIndex] = row.File;
                    arr[burstIndex] = row.Burst;
                    arr[startLabelIndex] = row.StartLabel;
                    arr[stopLabelIndex] = row.StopLabel;
                    arr[commentIndex] = row.Comment;
                    sb.AppendLine(string.Join("\t", arr));
                }
            }
            #endregion

            sw.Write(sb.ToString());
        }
    }
}
