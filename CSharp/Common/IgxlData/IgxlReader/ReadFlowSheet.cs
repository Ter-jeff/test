﻿using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlReader
{
    public class ReadFlowSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int EndRowIndex = 4;
        private const int StartColumnIndex = 2;

        public SubFlowSheet GetSheet(Stream stream, string sheetName)
        {
            var subFlowSheet = new SubFlowSheet(sheetName);
            var isBackup = false;
            var i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > EndRowIndex)
                    {
                        var flowRow = GetFlowRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(flowRow.OpCode))
                        {
                            isBackup = true;
                            continue;
                        }

                        flowRow.IsBackup = isBackup;
                        subFlowSheet.AddRow(flowRow);
                    }
                    i++;
                }
            }
            return subFlowSheet;
        }

        private FlowRow GetFlowRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var flowRow = new FlowRow();
            flowRow.RowNum = row;
            flowRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            flowRow.ColumnA = content;
            content = GetCellText(arr, index);
            flowRow.Label = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Enable = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Job = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Part = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Env = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.OpCode = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Parameter = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.TName = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.TNum = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.LoLim = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.HiLim = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Scale = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Units = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Format = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.BinPass = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.BinFail = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.SortPass = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.SortFail = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Result = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.PassAction = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.FailAction = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.State = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.GroupSpecifier = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.GroupSense = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.GroupCondition = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.GroupName = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.DeviceSense = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.DeviceCondition = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.DeviceName = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.DebugAsume = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.DebugSites = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.CtProfileDataElapsedTime = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.CtProfileDataBackgroundType = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.CtProfileDataSerialize = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.CtProfileDataResourceLock = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.CtProfileDataFlowStepLocked = content;
            index++;
            content = GetCellText(arr, index);
            flowRow.Comment = content;
            index++;
            var comment1List = new List<string>();
            for (var i = index; i <= index + 1; i++)
                if (!string.IsNullOrEmpty(GetCellText(arr, i)))
                    comment1List.Add(GetCellText(arr, i));
                else
                    comment1List.Add("");
            flowRow.Comment1 = string.Join("\t", comment1List);
            return flowRow;
        }

        private FlowRow GetFlowRow(ExcelWorksheet sheet, int row)
        {
            var flowRow = new FlowRow();
            flowRow.SheetName = sheet.Name;
            flowRow.RowNum = row;
            var content = GetCellText(sheet, row, 1);
            flowRow.ColumnA = content;
            var index = StartColumnIndex;
            content = GetCellText(sheet, row, index);
            flowRow.Label = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Enable = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Job = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Part = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Env = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.OpCode = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Parameter = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.TName = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.TNum = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.LoLim = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.HiLim = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Scale = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Units = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Format = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.BinPass = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.BinFail = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.SortPass = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.SortFail = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Result = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.PassAction = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.FailAction = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.State = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupSpecifier = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupSense = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupCondition = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupName = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DeviceSense = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DeviceCondition = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DeviceName = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DebugAsume = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DebugSites = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataElapsedTime = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataBackgroundType = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataSerialize = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataResourceLock = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataFlowStepLocked = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Comment = content;
            index++;
            var comment1List = new List<string>();
            for (var i = index; i <= sheet.Dimension.Columns; i++)
                if (!string.IsNullOrEmpty(GetCellText(sheet, row, index)))
                    comment1List.Add(GetCellText(sheet, row, index));
            flowRow.Comment1 = string.Join("\t", comment1List);
            return flowRow;
        }

        public List<string> GetEnables(Stream stream, string sheetName)
        {
            var enables = new List<string>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (line != null)
                    {
                        var arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
                        if (arr.Count() > 3)
                        {
                            var enable = arr[2];
                            if (!string.IsNullOrEmpty(enable))
                                enables.Add(enable);
                        }
                    }
                }
            }

            return enables.Distinct().ToList();
        }

        public SubFlowSheet GetSheet(string fileName, bool isRemoveBackup = false)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName), isRemoveBackup);
        }

        public SubFlowSheet GetSheet(ExcelWorksheet sheet, bool isRemoveBackup = false)
        {
            var subFlowSheet = new SubFlowSheet(sheet);
            var maxRowCount = sheet.Dimension.End.Row;

            // Set Row
            for (var i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                var flowRow = GetFlowRow(sheet, i);
                if (isRemoveBackup && string.IsNullOrEmpty(flowRow.OpCode))
                    break;
                subFlowSheet.AddRow(flowRow);
            }

            return subFlowSheet;
        }
    }
}