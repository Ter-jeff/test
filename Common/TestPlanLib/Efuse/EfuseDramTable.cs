using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;

namespace TestPlanLib.Efuse
{
    public partial class EfuseDramTable(string inPath, string stdf)
    {
        [GeneratedRegex("Category", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("^Fuse*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex1 = MyRegex1();

        public string InPath = inPath;
        public string StdfFile = stdf;
        public List<string> Titles = [];
        public List<List<string>> Rows = [];

        public void Parse()
        {
            try
            {
                //init
                Titles.Clear();
                Rows.Clear();

                if (Path.GetExtension(InPath).EqualsIgnoreCase(".txt"))
                {
                    ReadBdfTxt(InPath);
                }
                else
                {
                    ReadBdFxls(InPath);
                }
            }
            catch (Exception)
            {
                throw new Exception();
            }
        }

        private void ReadBdFxls(string path)
        {
            string line;
            using FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage ep = new ExcelPackage(fs);
            ExcelWorksheet ws = ep.Workbook.Worksheets["DRAM_CONFIG"] ?? ep.Workbook.Worksheets["DRAM_Table"];
            int startrow = 1;
            int endrow = ws.Dimension.End.Row;
            while (startrow != endrow)
            {
                line = ws.GetCellLine(startrow);
                startrow++;
                if (_regex.IsMatch(line))
                {
                    ReadHeader(ws);
                    break;
                }
            }

            try
            {
                while (startrow <= endrow)
                {
                    line = ws.GetCellLine(startrow);
                    if (_regex1.IsMatch(line))
                    {
                        List<string> vLine = [.. line.Split('\t')];
                        Rows.Add(vLine);
                    }
                    startrow++;
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private void ReadBdfTxt(string path)
        {
            string? line;
            try
            {
                using StreamReader sr = new StreamReader(path);
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();
                    if (_regex.IsMatch(line!))
                    {
                        ReadHeader(line!);
                        break;
                    }
                }

                sr.DiscardBufferedData();
                sr.BaseStream.Seek(0, SeekOrigin.Begin);

                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine()!.ToUpper();
                    if (_regex1.IsMatch(line))
                    {
                        List<string> vLine = [.. line.Split('\t')];
                        Rows.Add(vLine);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private void ReadHeader(ExcelWorksheet excelWorksheet)
        {
            int endCol = excelWorksheet.Dimension.End.Column;
            //int ebdRow = sheet.Dimension.End.Row;   
            for (int col = 1; col <= endCol; col++)
            {
                string header = excelWorksheet.Cells[1, col].Text.Trim();
                Titles.Add(header);
            }
        }

        private void ReadHeader(string line)
        {
            int endCol = line.Split('\t').Length;
            for (int col = 0; col < endCol; col++)
            {
                string header = line.Split('\t')[col].Trim();
                Titles.Add(header);
            }
        }
    }
}
