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
    public class InstanceSheet : IgxlSheet<InstanceRow>
    {
        public override string SheetType => "DTTestInstancesSheet";
        public override string IgxlSheetName => IgxlSheetNames.TestInstances;

        private readonly bool _isAllowOverwrite;

        public InstanceSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public InstanceSheet(string sheetName, string sourceSheet = "", bool isAllowOverwrite = false) : base(sheetName)
        {
            SourceInfo.Name = sourceSheet;
            _isAllowOverwrite = isAllowOverwrite;
        }

        public InstanceSheet(InstanceSheet instanceSheet) : base(instanceSheet)
        {
            if (instanceSheet == null)
            {
                return;
            }

            _isAllowOverwrite = instanceSheet._isAllowOverwrite;
            Rows = [.. instanceSheet.Rows.Select(x => x.Copy())];
        }

        public InstanceSheet Copy()
        {
            return new InstanceSheet(this);
        }

        public override void AddRow(InstanceRow instanceRow)
        {
            if (_isAllowOverwrite)
            {
                int tarInst = Rows.FindIndex(p => p.TestName == instanceRow.TestName);
                if (tarInst != -1)
                {
                    Rows.RemoveAt(tarInst);
                    Rows.Insert(tarInst, instanceRow);
                }
                else
                {
                    Rows.Add(instanceRow);
                }
            }
            else
            {
                Rows.Add(instanceRow);
            }
        }

        public override void Write(string fileName, string version = "")
        {
            Write2P4(fileName, version);
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
            int testNameIndex = GetIndexFrom(sheetInfo, "Test Name");
            int typeIndex = GetIndexFrom(sheetInfo, "Test Procedure", "Type");
            int nameIndex = GetIndexFrom(sheetInfo, "Test Procedure", "Name");
            int calledAsIndex = GetIndexFrom(sheetInfo, "Test Procedure", "Called As");
            int dcSpecsCategoryIndex = GetIndexFrom(sheetInfo, "DC Specs", "Category");
            int dcSpecsSelectorIndex = GetIndexFrom(sheetInfo, "DC Specs", "Selector");
            int acSpecsCategoryIndex = GetIndexFrom(sheetInfo, "AC Specs", "Category");
            int acSpecsSelectorIndex = GetIndexFrom(sheetInfo, "AC Specs", "Selector");
            int timeSetsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Time Sets");
            int edgeSetsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Edge Sets");
            int pinLevelsIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Pin Levels");
            int mixedSignalTimingIndex = GetIndexFrom(sheetInfo, "Sheet Parameters", "Mixed Signal Timing");
            int overlayIndex = GetIndexFrom(sheetInfo, "Overlay");
            int arg0Index = GetIndexFrom(sheetInfo, "Other Parameters", "Arg0");
            int arg1Index = GetIndexFrom(sheetInfo, "Other Parameters", "Arg1");
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
            var mainList = Rows.Where(x => !x.IsBackup).ToList();
            var backupList = Rows.Where(x => x.IsBackup).ToList();
            if (backupList.Count != 0)
            {
                PrintBackupEmptyRows(mainList);
                mainList.AddRange(backupList);
            }

            for (int index = 0; index < mainList.Count; index++)
            {
                InstanceRow row = mainList[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.TestName))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[testNameIndex] = row.TestName;
                    arr[typeIndex] = row.VbtType;
                    arr[nameIndex] = row.VbtName;
                    arr[calledAsIndex] = row.CalledAs;
                    arr[dcSpecsCategoryIndex] = row.DcCategory;
                    arr[dcSpecsSelectorIndex] = row.DcSelector;
                    arr[acSpecsCategoryIndex] = row.AcCategory;
                    arr[acSpecsSelectorIndex] = row.AcSelector;
                    arr[timeSetsIndex] = row.TimeSets;
                    arr[edgeSetsIndex] = row.EdgeSets;
                    arr[pinLevelsIndex] = row.PinLevels;
                    arr[mixedSignalTimingIndex] = row.MixedSignalTiming;
                    arr[overlayIndex] = row.Overlay;
                    arr[arg0Index] = row.ArgList;
                    for (int i = 0; i < row.Args.Count; i++)
                    {
                        string arg = row.Args[i];
                        arr[arg1Index + i] = arg;
                    }
                    arr[commentIndex] = row.Comment;
                }
                else
                {
                    arr = ["\t"];
                }
                sb.AppendLine(string.Join("\t", arr));
            }
            #endregion

            sw.Write(sb.ToString());
        }

        public List<Tuple<string, string>> GetUsedDcSpecs()
        {
            var list = new List<Tuple<string, string>>();
            IEnumerable<IGrouping<string, InstanceRow>> groups = Rows.Where(x => !string.IsNullOrEmpty(x.DcCategory) && !string.IsNullOrEmpty(x.DcSelector))
                    .GroupBy(x => x.DcCategory + "&" + x.DcSelector);
            foreach (IGrouping<string, InstanceRow> group in groups)
            {
                InstanceRow first = group.First();
                list.Add(Tuple.Create(first.DcCategory, first.DcSelector));
            }
            return list;
        }

        public void RemoveDuplicateInstance(bool includeSheetName = true)
        {
            var removeList = new List<int>();
            var groups = Rows.GroupBy(x => x.TestName.ToLower()).ToList();
            foreach (IGrouping<string, InstanceRow> group in groups)
            {
                if (group.Count() > 1)
                {
                    for (int i = 1; i < group.Count(); i++)
                    {
                        InstanceRow item = group.ElementAt(i);
                        if (group.First().GetDifferences(item, includeSheetName).Count == 0)
                        {
                            int index = Rows.FindIndex(x => x == item);
                            removeList.Add(index);
                        }
                    }
                }
            }

            removeList = [.. removeList.OrderByDescending(x => x)];
            foreach (int item in removeList)
            {
                Rows.RemoveAt(item);
            }
        }

        public void AddFooter(string block)
        {
            var footer = new InstanceRow
            {
                TestName = block + "_Footer_1",
                VbtType = "VBT",
                ArgList = "PrintInfo",
                VbtName = "Print_Footer"
            };
            footer.Args.Add(block);
            Rows.Add(footer);
        }

        public void AddHeader(string block)
        {
            var header = new InstanceRow
            {
                TestName = block + "_Header_1",
                VbtType = "VBT",
                ArgList = "PrintInfo",
                VbtName = "Print_Header"
            };
            header.Args.Add(block);
            Rows.Add(header);
        }

        public void AddHeaderFooter(string sheetName)
        {
            string block = sheetName.ReplaceStartsWith("Flow_", "");

            AddHeader(block);

            AddFooter(block);
        }
    }
}
