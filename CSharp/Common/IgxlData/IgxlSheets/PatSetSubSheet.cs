﻿using IgxlData.IgxlBase;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class PatSetSubSheet : IgxlSheet
    {
        #region Field

        private const string SheetType = "DTPatternSubroutineSheet";

        #endregion

        #region Property

        public List<PatSetSubRow> PatSetSubRows { set; get; }

        #endregion

        #region Constructor

        public PatSetSubSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            PatSetSubRows = new List<PatSetSubRow>();
            IgxlSheetName = IgxlSheetNameList.PatternSubroutine;
        }

        public PatSetSubSheet(string sheetName)
            : base(sheetName)
        {
            PatSetSubRows = new List<PatSetSubRow>();
            IgxlSheetName = IgxlSheetNameList.PatternSubroutine;
        }

        #endregion

        #region Member Function

        public void AddRow(PatSetSubRow igxlItem)
        {
            PatSetSubRows.Add(igxlItem);
        }

        // 20161027 add by JN 
        public long GetPatSetSubCnt()
        {
            return PatSetSubRows.Count;
        }
        // 20161027 add by JN 

        protected void WriteHeader()
        {
            const string header =
                "DTPatternSubroutineSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1	Pattern Subroutine";
            IgxlWriter.WriteLine(header);
            IgxlWriter.WriteLine();
            IgxlWriter.WriteLine();
        }

        protected void WriteColumnsHeader()
        {
        }

        protected void WriteRows()
        {
            foreach (var patSetSub in PatSetSubRows)
            {
                var row = new StringBuilder();
                row.Append("\t");
                row.Append(patSetSub.PatternFileName);
                row.Append("\t");
                row.Append(patSetSub.Comment);
                IgxlWriter.WriteLine(row.ToString());
            }
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.0";
            if (version == "2.0")
            {
                GetStreamWriter(fileName);
                WriteHeader();
                WriteColumnsHeader();
                WriteRows();
                CloseStreamWriter();
            }
            else
            {
                throw new Exception(string.Format("The PatternSubroutine sheet version:{0} is not supported!",
                    version));
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (PatSetSubRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var patternFilenameIndex = GetIndexFrom(igxlSheetsVersion, "Pattern Filename");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
                             IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < PatSetSubRows.Count; index++)
                {
                    var row = PatSetSubRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.PatternFileName))
                    {
                        arr[0] = row.ColumnA;
                        arr[patternFilenameIndex] = row.PatternFileName;
                        arr[commentIndex] = row.Comment;
                    }
                    else
                    {
                        arr = new[] { "\t" };
                    }

                    sw.WriteLine(string.Join("\t", arr));
                }

                #endregion
            }
        }

        #endregion

        public void AddRows(List<PatSetSubRow> patSetSubRows)
        {
            foreach (var patSetSubRow in patSetSubRows)
            {
                if (!PatSetSubRows.Exists(x => x.PatternFileName.Equals(patSetSubRow.PatternFileName, StringComparison.CurrentCultureIgnoreCase)))
                    AddRow(patSetSubRow);
            }
        }

    }
}