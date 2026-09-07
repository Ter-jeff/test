using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public partial class IdsDistributionReader
    {
        private const string RegexPerformance = "(?<pmode>M[a-zA-Z0-9]{4}[a-zA-Z0-9]?)";
        [GeneratedRegex(RegexPerformance, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]

        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regexPolation = new Regex(@"^(extra|inter)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public static IdsDistributionTable Read(string inPath)
        {
            var oIdsDisTable = new IdsDistributionTable();
            //init
            oIdsDisTable.AllIdsPowers.Clear();

            var allLines = new List<string>();
            using (var sr = new StreamReader(inPath))
            {
                //STEP1. Get all line, and remove empty line
                string? line;
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();
                    if (line?.Length == 0)
                    {
                        continue;
                    }

                    allLines.Add(line!);
                }

                int lineIdx = 0;
                int instanceCount = 0;
                string[]? instTypeAry = null;
                //Line#2. Module Type
                //TD		MBIST		SPI		TMPS		LDCBFD		
                for (; lineIdx < allLines.Count; lineIdx++)
                {
                    line = allLines[lineIdx];
                    if (line.Contains("TD") && line.Contains("MBIST"))  //TD		MBIST		SPI		TMPS		LDCBFD
                    {
                        instTypeAry = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                        instanceCount = instTypeAry.Length;
                        break;
                    }
                }
                lineIdx++;
                allLines.RemoveRange(0, lineIdx);

                //iteration read all ids range
                var oneBlock = new List<string>();
                while (GetOneBlock(ref allLines, ref oneBlock))
                {
                    var objIdsPower = new IdsPower();
                    var allIdsRngs = new IdsInfo[instanceCount];
                    for (int i = 0; i < instanceCount; i++)
                    {
                        allIdsRngs[i] = new IdsInfo();
                        if (instTypeAry != null)
                        {
                            allIdsRngs[i].TypeName = instTypeAry[i];
                        }
                    }

                    //Line#3:  MC607		MC607		MC607		MC607		MC607
                    lineIdx = 0;
                    line = oneBlock[lineIdx++];
                    //                        line = sr.ReadLine();
                    string[] spt1 = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                    objIdsPower.PowerName = spt1[0];

                    //Line#4:  IDS Range	Start Bin	IDS Range	Start Bin	IDS Range	Start Bin	IDS Range	Start Bin	IDS Range	Start Bin	
                    line = oneBlock[lineIdx++];

                    //oIdsDisTable.allIdsPowers[pwrIdx].allIdsRngs[TD_IDX].idsRng;
                    //oIdsDisTable.allIdsPowers[pwrIdx].allIdsRngs[SPI_IDX].startBin;
                    //Line#5: 0	4	0	4	0	4	0	4	0	4	
                    //iterate find all the IdsRng for one performance power
                    while (lineIdx < oneBlock.Count)
                    {
                        line = oneBlock[lineIdx++];
                        //                            line = sr.ReadLine();
                        //if (line.Contains("End") || line.Contains("END") || line.Contains("end"))
                        //   break;
                        string[] spt = line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries);
                        if (spt.Length != instanceCount * 2)
                        {
                            break;
                        }

                        if (spt.Length != 0 && spt[0].EqualsIgnoreCase("END"))
                        {
                            //isParserEnd = true;
                            break;
                        }

                        for (int i = 0; i < instanceCount; i++)
                        {
                            allIdsRngs[i].IdsRng.Add(double.Parse(spt[0 + (2 * i)]));
                            allIdsRngs[i].StartBin.Add(int.Parse(spt[1 + (2 * i)]));
                        }
                    }
                    //Line#6: 121.68	0	121.68	0	121.68	0	121.68	0	121.68	0	

                    //assign struct value
                    for (int i = 0; i < instanceCount; i++)
                    {
                        objIdsPower.IdsInfos.Add(allIdsRngs[i]);
                    }

                    oIdsDisTable.AllIdsPowers.Add(objIdsPower);
                }
            }
            return oIdsDisTable;
        }

        public static IdsDistributionTable ReadStartEqn(string inPath)
        {
            var oIdsDisTable = new IdsDistributionTable();
            //init
            oIdsDisTable.AllIdsPowers.Clear();

            using (var sr = new StreamReader(inPath))
            {
                //STEP1. Get all line, and remove empty line
                string? line;
                while (!sr.EndOfStream)
                {
                    var objIdsPower = new IdsPower();
                    var allIdsRngs = new IdsInfo();
                    line = sr.ReadLine();
                    if (line?.Length == 0)
                    {
                        continue;
                    }

                    string[] spt1 = line!.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                    if (_regex.IsMatch(spt1[0]) && int.Parse(spt1[1]) > 0) // Ignore start EQN number <= 0
                    {
                        objIdsPower.PowerName = spt1[0];
                        /* For aligning with the original data format in IdsDistribution */
                        allIdsRngs.IdsRng.Add(0);
                        allIdsRngs.IdsRng.Add(10000);
                        /* For aligning with the original data format in IdsDistribution */
                        allIdsRngs.StartBin.Add(int.Parse(spt1[1]));
                        objIdsPower.IdsInfos.Add(allIdsRngs);
                        oIdsDisTable.AllIdsPowers.Add(objIdsPower);

                    }
                }
            }
            return oIdsDisTable;
        }

        public static IdsDistributionTable ReadStartEqnSheet(string filePath, string? job = null)
        {
            var oIdsDisTable = new IdsDistributionTable();
            //init
            oIdsDisTable.AllIdsPowers.Clear();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets["Start_EQN"];
                if (ws == null)
                {
                    Response.Report("Start_EQN sheet not found");
                }

                if (ws?.Dimension is not { })
                {
                    Response.Report("Start_EQN sheet is empty");
                    return oIdsDisTable;
                }

                int rowCount = 0;
                int colCount = 0;
                if (ws.Dimension != null)
                {
                    rowCount = ws.Dimension.End.Row;
                    colCount = ws.Dimension.End.Column;
                }

                for (int r = 1; r <= rowCount; r++)
                {
                    var objIdsPower = new IdsPower();
                    var allIdsRngs = new IdsInfo();
                    List<string> cellValues = new List<string>();

                    for (int c = 1; c <= colCount; c++)
                    {
                        string text = ws.Cells[r, c].Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            cellValues.Add(text);
                        }
                    }

                    if (cellValues.Count == 0)
                    {
                        continue;
                    }

                    string[] spt1 = cellValues.ToArray();
                    if (spt1.Length < 2)
                    {
                        continue;
                    }

                    if (_regex.IsMatch(spt1[0]) && int.TryParse(spt1[1], out int bin) && bin > 0)
                    {
                        objIdsPower.PowerName = spt1[0];
                        allIdsRngs.IdsRng.Add(0);
                        allIdsRngs.IdsRng.Add(10000);
                        allIdsRngs.StartBin.Add(bin);
                        objIdsPower.IdsInfos.Add(allIdsRngs);
                        if (job != null)
                        {
                            objIdsPower.Job = job;
                        }
                        oIdsDisTable.AllIdsPowers.Add(objIdsPower);
                    }
                    else if (_regexPolation.IsMatch(spt1[1]))
                    {
                        objIdsPower.PowerName = spt1[0];
                        allIdsRngs.Polation.Add(spt1[1].Trim());
                        objIdsPower.IdsInfos.Add(allIdsRngs);
                        if (job != null)
                        {
                            objIdsPower.Job = job;
                        }
                        oIdsDisTable.AllIdsPowers.Add(objIdsPower);
                    }
                }
            }
            return oIdsDisTable;
        }

        private static bool GetOneBlock(ref List<string> allLines, ref List<string> oneBlock)
        {
            //init

            string line = "";
            int lineIdx = 0;
            oneBlock.Clear();

            if (allLines.Count == 0)
            {
                return false;
            }

            //Get start
            for (; lineIdx < allLines.Count; lineIdx++)
            {
                line = allLines[lineIdx];
                string[] spt1 = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                if (spt1.Length == 5 && _regex.IsMatch(spt1[0]))
                {
                    break;
                }
            }
            oneBlock.Add(line);

            //Read until eof or next block
            lineIdx++;
            for (; lineIdx < allLines.Count; lineIdx++)
            {
                line = allLines[lineIdx];

                string[] spt1 = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                if (spt1.Length == 5 && _regex.IsMatch(spt1[0]))
                {
                    break;
                }

                oneBlock.Add(line);
            }

            allLines.RemoveRange(0, lineIdx);

            return true;
        }
    }

    public class IdsInfo
    {
        //TD/MBIST/SPI....
        public string TypeName = "";
        public List<double> IdsRng = [];
        public List<int> StartBin = [];
        public List<string> Polation = [];
    }

    public class IdsPower
    {
        //Notice, I just decided to hard code Module(td/MBist/SPI) idx here without using compare to search
        //once IDS_DISTRIBUTION Table changed it's field, the hard code value must change too.
        public const int TdIdx = 0;
        public const int MbistIdx = 1;
        public const int SpiIdx = 2;
        public const int TmpsIdx = 3;
        public const int LdcbfdIdx = 4;
        public string Job = "";
        public string PowerName = "";
        public List<IdsInfo> IdsInfos = [];
    }

    public class IdsDistributionTable
    {
        public List<IdsPower> AllIdsPowers = [];
    }
}
