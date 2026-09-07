using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.Wireless.DVDC.InputReader;
using Automation.Static;

using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public partial class InputReader
    {
        [GeneratedRegex("HardIP_DC", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        public List<PatternRow> PlanItems = [];

        /// <summary>
        /// Read Inputs
        /// </summary>
        /// <param name="patInfoFile">Get Meas information</param>
        /// <param name="scghFile">Get Block information</param>
        /// <param name="patListFile"></param>
        public void ReadInput()
        {
            string outString;

            if (File.Exists(LocalSpecs.HardIpInfoFileName))
            {
                outString = string.Format("Reading PatInfo file ...");
                Response.Report(outString, percentage: Convert.ToInt32(30));
                var patInfoReader = new PatInfoReader();
                List<HardIpInfo> patInfo = patInfoReader.ExtractHardIpInfos(LocalSpecs.HardIpInfoFileName);
                LocalSpecs.HardIpInfos = new HardIpInfos(patInfo);
                //foreach (KeyValuePair<string, List<HardIpInfo>> info in LocalSpecs.HardIpInfos)
                //{
                //    info.UseThisVersion = true;
                //}
            }
            else
            {
                throw new Exception("Missing patInfoFile ...");
            }
            if (!File.Exists(LocalSpecs.ScghFileName) && !File.Exists(LocalSpecs.PatternListCsvFileName))
            {
                throw new Exception("Missing scgh file and pattern list file");
            }

            if (File.Exists(LocalSpecs.ScghFileName))
            {
                outString = string.Format("Reading SCGH file ...");
                Response.Report(outString, percentage: Convert.ToInt32(40));
            }

            if (File.Exists(LocalSpecs.TestPlanFileName))
            {
                Response.Report(string.Format("Reading Test Plan file ..."), percentage: 60);
                ExcelWorkbook planWBook = EpWorkbook.TestPlanWorkbook;

                foreach (ExcelWorksheet sheet in planWBook.Worksheets)
                {
                    TestPlanReader? reader = null;
                    if (MyRegex().IsMatch(sheet.Name))
                    {
                        continue;
                    }

                    if (sheet.Name.StartsWithIgnoreCase(NeededSheets.PrefixWireless) ||
                        sheet.Name.StartsWithIgnoreCase(NeededSheets.PrefixLcd))
                    {
                        reader = new WirelessTestPlanReader();
                    }
                    else if (sheet.Name.StartsWithIgnoreCase(NeededSheets.PrefixHardIp))
                    {
                        reader = new TestPlanReader();
                    }

                    if (reader != null)
                    {
                        Response.Report(string.Format("Reading Test Plan sheet ... {0}", sheet.Name), percentage: 80);
                        PlanItems.AddRange(reader.ReadSheet(sheet).PatternRows);
                    }
                }
            }
        }
    }
}
