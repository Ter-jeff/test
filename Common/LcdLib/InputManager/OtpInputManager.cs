using System.IO;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.InputManager;
using Automation.Static;

using CommonLib.Enums;

using LcdLib.InputManager.Data;
using LcdLib.OTP.Reader;

using LogLib.Static;

using OfficeOpenXml;

namespace LcdLib.InputManager
{
    public class OtpInputManager(ExcelWorkbook excelWorkbook) : InputManagerBase<OtpInputData>(excelWorkbook)
    {
        public override OtpInputData Read()
        {
            var result = new OtpInputData();
            if (LocalSpecs.ScghFileName != "N/A" && !string.IsNullOrEmpty(LocalSpecs.ScghFileName))
            {
                if (File.Exists(LocalSpecs.ScghFileName))
                {
                    Response.Report("Reading Otp Patterns in SCGH ...", EnumMessageLevel.General, 20);
                    var excel = new ExcelPackage(new FileInfo(LocalSpecs.ScghFileName));
                    ExcelWorkbook scghWorkbook = excel.Workbook;
                    ScghData efuseScghSheet = new ScghData().LoadEfuseFromHardIpBistScghData(scghWorkbook, true);
                    result.OtpPatternRows = OtpPatternReader.GetOtpPatternRows(efuseScghSheet);
                }
            }

            return result;
        }
    }
}
