using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;

namespace TestPlanLib.Efuse
{
    public class FuseCheckTable(string inPath, string stdf)
    {
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
            using ExcelPackage ep = new ExcelPackage(new FileInfo(path));
            ExcelWorksheet ws = ep.Workbook.Worksheets["FuseCheckTable"];
            int startrow = 1;
            int endrow = ws.Dimension.End.Row;
            while (startrow != endrow)
            {
                string line = ws.GetCellLine(startrow);
                startrow++;
                ReadHeader(ws);
                break;
            }

            try
            {
                while (startrow <= endrow)
                {
                    string line = ws.GetCellLine(startrow);
                    List<string> arr = [.. line.Split('\t')];
                    Rows.Add(arr);
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
                //-----------------------------------------------------------
                line = sr.ReadLine();
                ReadHeader(line!);
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine()!.ToUpper();
                    List<string> vLine = [.. line.Split('\t')];
                    Rows.Add(vLine);

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
