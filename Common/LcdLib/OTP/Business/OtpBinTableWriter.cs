using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LcdLib.Const;

using TestPlanLib.BinNumberLegacy;
using TestPlanLib.Singleton;

namespace LcdLib.OTP.Business
{
    internal class OtpBinTableWriter
    {
        public static void GenBinTable(List<string> existOtpFlags)
        {
            BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            List<BinTableRow> otpBinTables = GenerateBinTableRows(existOtpFlags);
            binTable.AddRows(otpBinTables);

            // MainInitEnableWd
            var flags = otpBinTables.SelectMany(x => x.ItemList.Split([','], StringSplitOptions.RemoveEmptyEntries)).Distinct(StringExtensions.IgnoreCase).ToList();
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, EFuseConst.Efuse, FolderStructure.DirMain);
        }

        private static List<BinTableRow> GenerateBinTableRows(List<string> existOtpFlags)
        {
            var binTableRows = new List<BinTableRow>();
            BinNumberSingletonLegacy.Instance().SetStartBinNumber(90);
            //var blankOrder = new List<string>()
            //{
            //    BankType.Ecid,
            //    BankType.Cfg,
            //    BankType.UdrP,
            //    BankType.UdrP0,
            //    BankType.UdrP1,
            //    BankType.UdrE,
            //    BankType.UdrE0,
            //    BankType.UdrE1,
            //    BankType.Mon,
            //    BankType.Unknow
            //};
            //var groups = existOtpFlags.GroupBy(x => EFuseConst.GetBankName(x)).OrderBy(y => blankOrder.IndexOf(y.Key));
            foreach (string flag in existOtpFlags)
            {
                var binNumPara = new BinNumDefPara(EnumBinNumDefBlock.eFuse, OtpConst.Otp);

                BinNumberSingletonLegacy.Instance().GetBinNumDefRow(binNumPara, out BinNumDefRow bin);
                var row = new BinTableRow();
                string binTableName = flag.Replace("F_", "Bin_");

                row.Name = binTableName;
                row.ItemList = flag;
                row.Items.Add("T");
                row.Op = "OR";
                row.Result = "Fail";
                //row.Sort = BinNumberSingleton.Instance().GetEfuseSoftBinNumber(row);//bin.CurrentSoftBin.ToString();
                row.Bin = bin.HardBin;
                binTableRows.Add(row);
            }
            return binTableRows;
        }

        public static BinTableRow GenBinTableRow(string name)
        {
            string type = EFuseConst.GetBankName(name);

            if (EFuseConst.BankIsUdr(type))
            {
                type = "UDR";
            }
            else if (type.EqualsIgnoreCase(BankType.Unknow))
            {
                type = "Reserve";
            }

            var binNumPara = new BinNumDefPara(EnumBinNumDefBlock.eFuse, type);
            BinNumberSingletonLegacy.Instance().GetBinNumDefRow(binNumPara, out BinNumDefRow bin);
            var row = new BinTableRow { Name = name, ItemList = name.Replace("Bin_EFUSE_", "F_") };
            row.Items.Add("T");
            row.Op = "AND";
            row.Result = "Fail-stop";
            //bin.CurrentSoftBin.ToString();
            row.Sort = BinNumberSingletonLegacy.Instance().GetEfuseSoftBinNumber(row);
            row.Bin = bin.HardBin;
            return row;
        }
    }
}
