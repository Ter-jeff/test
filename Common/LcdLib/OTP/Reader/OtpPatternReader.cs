using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Singleton;

using CommonLib.Extension;

using ScghLib.Reader;

namespace LcdLib.OTP.Reader
{
    public class OtpPatternReader
    {
        public MultiTestSettingSheetsSingleton? MultiTestSettingSheetsSingleton;

        public static List<OtpPatternRow> GetOtpPatternRows(ScghData scghData)
        {
            var otpPatRows = new List<OtpPatternRow>();
            var otpScghRowList = new List<ProdCharSheetRow>();
            if (scghData.HardIpSheetRowList.Count != 0)
            {
                otpScghRowList.AddRange(scghData.HardIpSheetRowList);
            }

            if (scghData.ScanScghRowList.Count != 0)
            {
                otpScghRowList.AddRange(scghData.ScanScghRowList);
            }

            var otpWriteItems = otpScghRowList.Where(x => x.Block.EqualsIgnoreCase("OTP_WRITE")).ToList();
            var otpReadItems = otpScghRowList.Where(x => x.Block.EqualsIgnoreCase("OTP_READ")).ToList();
            var otpFunctionalItems = otpScghRowList.Where(x => x.Block.EqualsIgnoreCase("OTP")).ToList();

            //write
            if (otpWriteItems.Count != 0)
            {
                var bankList = new List<EnumOtpBankName> { EnumOtpBankName.Ecid, EnumOtpBankName.Crc, EnumOtpBankName.All };
                var writeTypeList = new List<EnumOtpWriteType> { EnumOtpWriteType.Multishot, EnumOtpWriteType.Multishot4Byte, EnumOtpWriteType.OneShot };
                foreach (ProdCharSheetRow data in otpWriteItems)
                {
                    foreach (EnumOtpBankName bankName in bankList)
                    {
                        foreach (EnumOtpWriteType writeType in writeTypeList)
                        {
                            var otpPatRow = new OtpPatternRow
                            {
                                Block = data.Block,
                                Mode = data.Mode,
                                Item = data.Item,
                                BankName = bankName,
                                OtpReadWrite = EnumOtpReadWrite.Write,
                                OtpWriteType = writeType,
                                InitList = data.InitList,
                                PayloadList = data.PayloadList,
                                PatList = []
                            };
                            otpPatRow.PatList.AddRange(data.InitList);
                            otpPatRow.PatList.AddRange(data.PayloadList);
                            otpPatRow.SheetName = data.SourceSheetName;
                            otpPatRow.RowNum = data.RowNum;
                            otpPatRows.Add(otpPatRow);
                        }
                    }
                }
            }

            //read
            if (otpReadItems.Count != 0)
            {
                var readTypeList = new List<EnumOtpReadType> { EnumOtpReadType.Multishot4Byte, EnumOtpReadType.OneShot8Bits, EnumOtpReadType.OneShot32Bits };
                foreach (ProdCharSheetRow data in otpReadItems)
                {
                    foreach (EnumOtpReadType readType in readTypeList)
                    {
                        var otpPatRow = new OtpPatternRow
                        {
                            Block = data.Block,
                            Mode = data.Mode,
                            Item = data.Item,
                            BankName = EnumOtpBankName.All,
                            OtpReadWrite = EnumOtpReadWrite.Read,
                            OtpReadType = readType,
                            InitList = data.InitList,
                            PayloadList = data.PayloadList,
                            PatList = []
                        };
                        otpPatRow.PatList.AddRange(data.InitList);
                        otpPatRow.PatList.AddRange(data.PayloadList);
                        otpPatRow.SheetName = data.SourceSheetName;
                        otpPatRow.RowNum = data.RowNum;
                        otpPatRows.Add(otpPatRow);
                    }
                }
            }

            //functional 
            foreach (ProdCharSheetRow data in otpFunctionalItems)
            {
                var otpPatRow = new OtpPatternRow
                {
                    Block = data.Block,
                    Mode = data.Mode,
                    Item = data.Item,
                    BankName = EnumOtpBankName.None,
                    OtpReadWrite = EnumOtpReadWrite.None,
                    InitList = data.InitList,
                    PayloadList = data.PayloadList,
                    PatList = []
                };
                otpPatRow.PatList.AddRange(data.InitList);
                otpPatRow.PatList.AddRange(data.PayloadList);
                otpPatRow.SheetName = data.SourceSheetName;
                otpPatRow.RowNum = data.RowNum;
                otpPatRows.Add(otpPatRow);
            }

            return otpPatRows;
        }
    }
}
