using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.ErrorReport.Base;

using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;
using IgxlLib.Utility;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    [DebuggerDisplay("{Name} ")]
    public abstract class IgxlSheet<T> : IIgxlSheet where T : IgxlRow
    {
        protected string? FilePath;
        protected string? IgxlSheetContext;
        public SourceInfo SourceInfo = new();

        public abstract string SheetType { get; }
        public abstract string IgxlSheetName { get; }
        public string Name { get; set; }
        public List<T> Rows { get; set; } = [];

        private readonly List<Error> _errors = [];

        protected IgxlSheet(ExcelWorksheet excelWorksheet)
        {
            IgxlSheetContext = excelWorksheet.Cells[1, 1].Text;
            Name = excelWorksheet.Name;
        }

        protected IgxlSheet(string sheetName)
        {
            Name = sheetName;
        }

        // Rows is intentionally left uncopied here (default []) — IgxlRow has no
        // common virtual Copy(), so each derived sheet class must deep-copy its own
        // concretely-typed Rows in its own copy constructor after calling base(...).
        protected IgxlSheet(IgxlSheet<T> ts) : this(ts?.Name ?? "")
        {
            if (ts == null)
            {
                return;
            }

            FilePath = ts.FilePath;
            IgxlSheetContext = ts.IgxlSheetContext;
            SourceInfo = ts.SourceInfo.Copy();
            _errors = ts._errors?.Select(row => row.Copy()).ToList() ?? [];
        }

        public virtual void AddRow(T t)
        {
            Rows.Add(t);
        }

        public void AddRows(List<T> ts)
        {
            Rows.AddRange(ts);
        }

        public void RemoveRow(T t)
        {
            Rows.Remove(t);
        }

        public void AddError(ErrorCode errorCode, string sheetName, int rowNum, int colNum, string[] messageArgs, ErrorInfo? errorInfo = null)
        {
            var error = new Error(errorCode, sheetName, rowNum, colNum, messageArgs, errorInfo);
            _errors.Add(error);
        }

        public List<Error> GetErrors()
        {
            return _errors;
        }

        public int InsertRow(int index, T t)
        {
            if (index == -1)
            {
                Rows.Add(t);
                return index;
            }

            Rows.Insert(index, t);
            return ++index;
        }

        public static Dictionary<string, Dictionary<string, SheetInfo>> GetIgxlSheetsVersion()
        {
            return IgxlSheetHelpers.GetIgxlSheetsVersion();
        }

        protected virtual void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            //nothing to do in base class
        }

        public abstract void Write(string fileName, string version = "");

        public virtual void Write2P0(string fileName, string version)
        {
            WriteByVersion(fileName, version, "2.0");
        }

        public virtual void Write2P1(string fileName, string version)
        {
            WriteByVersion(fileName, version, "2.1");
        }

        public virtual void Write2P2(string fileName, string version)
        {
            WriteByVersion(fileName, version, "2.2");
        }

        public virtual void Write2P3(string fileName, string version)
        {
            WriteByVersion(fileName, version, "2.3");
        }

        public virtual void Write2P4(string fileName, string version)
        {
            WriteByVersion(fileName, version, "2.4");
        }

        public virtual void Write2P5(string fileName, string version)
        {
            WriteByVersion(fileName, version, "2.5");
        }

        public virtual void Write3P1(string fileName, string version)
        {
            WriteByVersion(fileName, version, "3.1");
        }

        public void WriteByVersion(string fileName, string version, string defaultVersion)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? "");

            if (string.IsNullOrEmpty(version))
            {
                version = defaultVersion;
            }

            Dictionary<string, Dictionary<string, SheetInfo>> sheetClassMapping = GetIgxlSheetsVersion();
            if (!sheetClassMapping.TryGetValue(IgxlSheetName, out Dictionary<string, SheetInfo>? versionMap))
            {
                return;
            }

            if (versionMap.TryGetValue(version, out SheetInfo? sheetInfo) || versionMap.TryGetValue(defaultVersion, out sheetInfo))
            {
                WriteSheet(fileName, version, sheetInfo);
            }
        }

        protected static int GetIndexFrom(SheetInfo sheetInfo, string name)
        {
            return IgxlSheetHelpers.GetIndexFrom(sheetInfo, name);
        }

        protected static int GetIndexFrom(SheetInfo sheetInfo, string name, string subString)
        {
            return IgxlSheetHelpers.GetIndexFrom(sheetInfo, name, subString);
        }

        protected static int GetMaxCount(SheetInfo sheetInfo)
        {
            return IgxlSheetHelpers.GetMaxCount(sheetInfo);
        }

        protected void SetRelativeColumn(SheetInfo sheetInfo, int i, string[] arr, int relativeColumnIndex)
        {
            if (sheetInfo.Columns.RelativeColumn != null)
            {
                foreach (Column item in sheetInfo.Columns.RelativeColumn)
                {
                    if (item.indexFrom == item.indexTo && item.rowIndex == i)
                    {
                        arr[relativeColumnIndex - 1 + item.indexFrom] = item.columnName;
                    }
                }
            }
        }

        protected static void SetColumns(SheetInfo sheetInfo, int i, string[] arr)
        {
            IgxlSheetHelpers.SetColumns(sheetInfo, i, arr);
        }

        protected static void SetField(SheetInfo sheetInfo, int i, string[] arr)
        {
            IgxlSheetHelpers.SetField(sheetInfo, i, arr);
        }

        protected static void WriteHeader(string[] arr, StringBuilder stringBuilder)
        {
            IgxlSheetHelpers.WriteHeader(arr, stringBuilder);
        }

        protected string WriteHeader(string[] arr)
        {
            return IgxlSheetHelpers.WriteHeader(arr);
        }

        protected void PrintBackupEmptyRows(StringBuilder stringBuilder)
        {
            for (int i = 0; i < 10; i++)
            {
                stringBuilder.AppendLine("");
            }
        }

        protected void PrintBackupEmptyRows(List<T> ts)
        {
            ts.AddRange(Enumerable.Range(0, 10).Select(_ => Activator.CreateInstance<T>()));
        }

        protected StringBuilder WriteSheet(string version, SheetInfo sheetInfo, string sheetType)
        {
            var sb = new StringBuilder();

            if (Rows.Count == 0)
            {
                return sb;
            }

            int maxCount = GetMaxCount(sheetInfo);

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(sheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                SetField(sheetInfo, i, arr);

                SetColumns(sheetInfo, i, arr);

                WriteHeader(arr, sb);
            }
            #endregion

            WriteData(Rows, sb);
            return sb;
        }

        protected void WriteData(List<T> ts, StringBuilder stringBuilder)
        {
            var mainList = ts.Where(y => !y.IsBackup).ToList();
            var backupList = ts.Where(y => y.IsBackup).ToList();
            if (backupList.Count != 0)
            {
                PrintBackupEmptyRows(mainList);
                mainList.AddRange(backupList);
            }
            for (int index = 0; index < mainList.Count; index++)
            {
                T row = mainList[index];
                if (row is IIgxlRow igxlRow)
                {
                    string[] arr = igxlRow.Print();
                    stringBuilder.AppendLine(string.Join("\t", arr));
                }
            }
        }

        protected StringBuilder WriteContent2P0(string version)
        {
            var sb = new StringBuilder();
            Dictionary<string, Dictionary<string, SheetInfo>> sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.TryGetValue(IgxlSheetName, out Dictionary<string, SheetInfo>? dic))
            {
                if (dic.TryGetValue(version, out SheetInfo? igxlSheetsVersion1))
                {
                    sb = WriteSheet(version, igxlSheetsVersion1, SheetType);
                }
                else if (dic.TryGetValue("2.0", out SheetInfo? igxlSheetsVersion))
                {
                    sb = WriteSheet(version, igxlSheetsVersion, SheetType);
                }
            }
            return sb;
        }

        protected void WriteByOneVersion(string fileName, StringBuilder? stringBuilder)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? string.Empty);

            if (stringBuilder != null)
            {
                using var sw = new StreamWriter(fileName, false);
                sw.Write(stringBuilder.ToString());
            }
        }
    }

    public class SourceInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Block { get; set; } = string.Empty;

        public SourceInfo(SourceInfo sourceInfo)
        {
            if (sourceInfo == null)
            {
                return;
            }

            Name = sourceInfo.Name;
            Block = sourceInfo.Block;
        }

        public SourceInfo()
        {
        }

        public SourceInfo Copy()
        {
            return new SourceInfo(this);
        }
    }
}
