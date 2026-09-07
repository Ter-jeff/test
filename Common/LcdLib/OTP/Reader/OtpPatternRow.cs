using System.Collections.Generic;

namespace LcdLib.OTP.Reader
{
    public class OtpPatternRow
    {
        public string SheetName = "";
        public int RowNum = 0;
        public string Block = "";
        public string Mode = "";
        public string Item = "";
        public EnumOtpBankName BankName = EnumOtpBankName.None;
        public EnumOtpReadWrite OtpReadWrite = EnumOtpReadWrite.None;
        public EnumOtpWriteType OtpWriteType = EnumOtpWriteType.None;
        public EnumOtpReadType OtpReadType = EnumOtpReadType.None;
        public List<string> InitList = [];
        public List<string> PayloadList = [];
        public List<string> PatList = [];
    }

    public enum EnumOtpReadWrite
    {
        Read, Write, None
    }

    public enum EnumOtpBankName
    {
        All, Ecid, Crc, None
    }

    public enum EnumOtpWriteType
    {
        Multishot, Multishot4Byte, OneShot, None
    }

    public enum EnumOtpReadType
    {
        Multishot4Byte, OneShot8Bits, OneShot32Bits, None
    }
}
