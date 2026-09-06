using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Enums;

using OfficeOpenXml;

using RfLib.Dvdc.Base;

namespace RfLib.Dvdc.GenVbt
{
    public partial class WirelessVBTGenerator
    {
        [GeneratedRegex("Class_RFPathCtrlUtil", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("RFPC_Relay_Table", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("Cal_Table", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();

        public static void WorkFlow(Dictionary<string, HardIpSheet> planDic, string testplanfile)
        {
            //search specified VBT First
            //Determine whether need to edit VBT
            //1. insert specified parameter with calc
            #region Generate Wireless RF Functional Test VBT

            string[] library = Directory.GetFiles(FolderStructure.DirLib, "*.*", SearchOption.AllDirectories);

            #region Overwrite Class_RFPathCtrlUtil.cls file

            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                string? rfPathCtrlUtilVbtFile = library.FirstOrDefault(_regex.IsMatch);

                //var samplename = new StreamWriter(RFPathCtrlUtilVBTFile);
                using var ep = new ExcelPackage(new FileInfo(testplanfile));
                ExcelWorksheet? relay = null;
                ExcelWorksheet? cal = null;
                foreach (ExcelWorksheet sheet in ep.Workbook.Worksheets)
                {
                    if (_regex2.IsMatch(sheet.Name))
                    {
                        relay = sheet;
                    }

                    if (_regex3.IsMatch(sheet.Name))
                    {
                        cal = sheet;
                    }
                }
                if (relay != null && cal != null)
                {
                    var relaytableinfo = new RelayTableReader();
                    relaytableinfo.Read(relay);
                    var caltableinfo = new CalTableReader();
                    caltableinfo.Read(cal);
                    var rfpCcontent = new List<string>();
                    RelayPathBusiness.Write(rfpCcontent, relaytableinfo.RelayPathItems, caltableinfo.CalPathToLut);
                    //ClassVbt.WriteVBFile(RFPathCtrlUtilVBTFile, RFPCcontent);
                }
            }

            #endregion

            #endregion

        }
    }
}
