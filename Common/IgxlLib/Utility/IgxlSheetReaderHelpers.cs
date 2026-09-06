using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.Utility
{
    public static class IgxlSheetReaderHelpers
    {
        private static string ReadFirstLine(string file)
        {
            using var sw = new StreamReader(file);
            return sw.ReadLine() ?? "";
        }

        public static List<string> GetSheetsByType(string path, EnumSheetType enumSheetType)
        {
            Dictionary<string, EnumSheetType> dic = GetSheetTypeDic(path);
            return [.. dic.ToList().Where(x => x.Value == enumSheetType).Select(x => x.Key)];
        }

        private static Dictionary<string, EnumSheetType> GetSheetTypeDic(string path)
        {
            return Directory.EnumerateFiles(path, "*.txt", SearchOption.AllDirectories).ToDictionary(file => file, file => IgxlLoaderHelpers.GetIgxlSheetType(ReadFirstLine(file)));
        }

        public static bool IsSheetType(EnumSheetType enumSheetType, string file)
        {
            return Regex.IsMatch(ReadFirstLine(file), enumSheetType.ToString(), RegexOptions.IgnoreCase);
        }

        public static ExcelWorksheet ConvertTxtToExcelSheet(string fileName)
        {
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            using FileStream stream = File.OpenRead(fileName);
            return ConvertStreamToExcelSheet(sheetName, stream);
        }

        public static ExcelWorksheet ConvertStreamToExcelSheet(string sheetName, Stream stream)
        {
            //can not "using"
            var excelPackage = new ExcelPackage();
            sheetName = sheetName.Replace("%20", " ");
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            int index = 0;
            using (var sr = new StreamReader(stream))
            {
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    index++;
                    string[] arr = line.Split(['\t'], StringSplitOptions.None);
                    for (int cnt = 0; cnt < arr.Length; cnt++)
                    {
                        sheet.Cells[index, 1 + cnt].Value = arr[cnt];
                    }
                }
            }
            return sheet;
        }

        public static ExcelWorksheet? GetExcelWorksheet(string igxl, string name)
        {
            using ZipArchive zipArchive = ZipFile.Open(igxl, ZipArchiveMode.Read);
            ZipArchiveEntry? entry = zipArchive.Entries.FirstOrDefault(e => Path.GetFileNameWithoutExtension(e.Name).EqualsIgnoreCase(name));
            return entry != null ? ConvertStreamToExcelSheet(name, entry.Open()) : null;
        }

        public static List<IIgxlSheet> GetIgxlSheetsByIgxlFile(string igxl, EnumSheetType enumSheetType)
        {
            using ZipArchive zipArchive = ZipFile.Open(igxl, ZipArchiveMode.Read);
            var entries = zipArchive.Entries.ToList();
            var igxlSheets = new List<IIgxlSheet>();
            foreach (ZipArchiveEntry zipArchiveEntry in entries)
            {
                string sheetName = Path.GetFileNameWithoutExtension(zipArchiveEntry.Name);
                EnumSheetType type;
                using (var sr = new StreamReader(zipArchiveEntry.Open()))
                {
                    string firstLine = sr.ReadLine() ?? "";
                    type = IgxlLoaderHelpers.GetIgxlSheetType(firstLine);
                }
                if (enumSheetType == type)
                {
                    using Stream stream = zipArchiveEntry.Open();
                    ExcelWorksheet sheet = ConvertStreamToExcelSheet(sheetName, stream);
                    igxlSheets.Add(GetIgxlSheet(sheet, enumSheetType));
                }
            }
            return igxlSheets;
        }

        public static IIgxlSheet CreateIgxlSheet(string file)
        {
            string firstLine = ReadFirstLine(file);
            EnumSheetType sheetType = IgxlLoaderHelpers.GetIgxlSheetType(firstLine);
            ExcelWorksheet sheet = ConvertTxtToExcelSheet(file);
            return GetIgxlSheet(sheet, sheetType);
        }

        private static IIgxlSheet GetIgxlSheet(ExcelWorksheet excelWorksheet, EnumSheetType enumSheetType)
        {
            var factory = new IgxlSheetFactory();
            return factory.CreateSheet(enumSheetType, excelWorksheet)!;
        }

        public static List<IIgxlSheet> GetIgxlSheets(string path, EnumSheetType enumSheetType)
        {
            return [.. Directory.GetFiles(path)
                .Where(file => Path.GetExtension(file).EqualsIgnoreCase(".txt"))
                .Where(file => enumSheetType == IgxlLoaderHelpers.GetIgxlSheetType(ReadFirstLine(file)))
                .Select(CreateIgxlSheet)];
        }
    }
}
