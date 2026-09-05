using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class BinTableGenerator
    {
        public static void Generate(List<CharPlanSheet> charPlanSheets)
        {
            BinTableSheet preGen = LocalSpecs.TestProgram.BintableSheets.FirstOrDefault
                (p => Regex.IsMatch(p.Name, "BinTable|Bin_Table", RegexOptions.IgnoreCase));
            if (preGen == null)
            {
                LocalSpecs.MessageWriter.WriteLine("[Error] Could not found Bin_Table or BinTable sheet in the prod progam!");
                return;
            }

            #region Bin_Char_Shmoo_Hole_Alarm_Allfail

            var binCharShmooHoleAlarmAllfail = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Hole_Alarm_Allfail",
                ItemList =
                    "F_Check_Shmoo_Hole_Ratio_Within_Spec,F_Check_Shmoo_Allfail_Ratio_Within_Spec,F_Check_Shmoo_Alarm_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9974",
                Bin = "5",
                Result = "Fail",
                Items = { "T", "T", "T" }
            };

            #endregion

            #region Bin_Char_Shmoo_Hole_Alarm

            var binCharShmooHoleAlarm = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Hole_Alarm",
                ItemList = "F_Check_Shmoo_Hole_Ratio_Within_Spec,F_Check_Shmoo_Alarm_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9975",
                Bin = "5",
                Result = "Fail",
                Items = { "T", "T" }
            };

            #endregion

            #region Bin_Char_Shmoo_Hole_Allfail

            var binCharShmooHoleAllfail = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Hole_Allfail",
                ItemList = "F_Check_Shmoo_Hole_Ratio_Within_Spec,F_Check_Shmoo_Allfail_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9976",
                Bin = "5",
                Result = "Fail",
                Items = { "T", "T" }
            };

            #endregion

            #region Bin_Char_Shmoo_Alarm_Allfail

            var binCharShmooAlarmAllfail = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Alarm_Allfail",
                ItemList = "F_Check_Shmoo_Allfail_Ratio_Within_Spec,F_Check_Shmoo_Alarm_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9977",
                Bin = "5",
                Result = "Fail",
                Items = { "T", "T" }
            };

            #endregion

            #region Bin_Char_Shmoo_Allfail

            var binCharShmooAllfail = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Allfail",
                ItemList = "F_Check_Shmoo_Allfail_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9978",
                Bin = "5",
                Result = "Fail",
                Items = { "T" }
            };

            #endregion

            #region Bin_Char_Shmoo_Alarm

            var binCharShmooAlarm = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Alarm",
                ItemList = "F_Check_Shmoo_Alarm_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9979",
                Bin = "5",
                Result = "Fail",
                Items = { "T" }
            };

            #endregion

            #region Bin_Char_Shmoo_Hole

            var binCharShmooHole = new BinTableRow
            {
                Name = "Bin_Char_Shmoo_Hole",
                ItemList = "F_Check_Shmoo_Hole_Ratio_Within_Spec",
                Op = "AND",
                Sort = "9980",
                Bin = "5",
                Result = "Fail",
                Items = { "T" }
            };

            #endregion

            #region Bin_Shmoo_Alarm

            var binShmooAlarm = new BinTableRow
            {
                Name = "Bin_initFlow_BinOut",
                ItemList = "F_initFlow_BinOut,F_Shmoo_Alarm",
                Op = "OR",
                Sort = "9988",
                Bin = "15",
                Result = "Fail-stop",
                Items = { "T", "T" }
            };

            #endregion

            var bintable = new BinTableSheet("Bin_Table");
            bool isNeedGenHardCodeBinRow = true;
            foreach (BinTableRow prebin in preGen.Rows)
            {
                if (isNeedGenHardCodeBinRow)
                {
                    isNeedGenHardCodeBinRow = false;
                    // add hard coded shmoo abnormal bin out bin table
                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooHoleAlarmAllfail.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooHoleAlarmAllfail);
                    }

                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooHoleAlarm.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooHoleAlarm);
                    }

                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooHoleAllfail.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooHoleAllfail);
                    }

                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooAlarmAllfail.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooAlarmAllfail);
                    }

                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooAllfail.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooAllfail);
                    }

                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooAlarm.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooAlarm);
                    }

                    if (!bintable.Rows.Exists(x => x.Name.Equals(binCharShmooHole.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        bintable.AddRow(binCharShmooHole);
                    }
                }

                if (IsAdd(prebin, binCharShmooHoleAlarmAllfail) &&
                    IsAdd(prebin, binCharShmooHoleAlarm) &&
                    IsAdd(prebin, binCharShmooHoleAllfail) &&
                    IsAdd(prebin, binCharShmooAlarmAllfail) &&
                    IsAdd(prebin, binCharShmooAllfail) &&
                    IsAdd(prebin, binCharShmooAlarm) &&
                    IsAdd(prebin, binCharShmooHole))
                {
                    bintable.AddRow(prebin);
                }
            }
            AddOrInsert(bintable, binShmooAlarm);

            var binoutFlags = charPlanSheets.SelectMany(p => p.CharList).Where(p => !string.IsNullOrEmpty(p.FailFlag)).Select(p => p.FailFlag).Distinct().ToList();

            foreach (string binoutFlag in binoutFlags)
            {
                string bintablename = Regex.Replace(binoutFlag, "^F_", "Bin_", RegexOptions.IgnoreCase);
                var binoutRow = new BinTableRow
                {
                    Name = bintablename,
                    ItemList =
                        binoutFlag,
                    Op = "AND",
                    Sort = "0",
                    Bin = "0",
                    Result = "Fail",
                    Items = { "T" }
                };
                bintable.AddRow(binoutRow);
            }

            // export bin table whose path is decided the GenTexOnly flag
            string outputFolder = Path.Combine(LocalSpecs.OutputFolder, ConstData.Binfolder);
            if (LocalSpecs.InputParam.GenTxtOnly)
            {
                outputFolder = LocalSpecs.OutputFolder;
            }

            string mainBinFile = Path.Combine(outputFolder, "Bin_Table" + ".txt");
            bintable.Write(mainBinFile);
            LocalSpecs.GenSheets.Add(bintable);
        }

        //This funciton check whether the bintable row is needed to add
        private static bool IsAdd(BinTableRow bin1, BinTableRow bin2)
        {
            if (bin1.Sort != bin2.Sort)
            {
                return true;
            }

            bool itemFlag = false;
            if (bin1.Items.Count == bin2.Items.Count)
            {
                for (int i = 0; i < bin1.Items.Count; i++)
                {
                    itemFlag = bin1.Items[i] == bin2.Items[i];
                }

            }
            if (!(bin1.Op == bin2.Op && bin1.Bin == bin2.Bin && bin1.Result == bin2.Result &&
                  itemFlag))
            {
                GeneralFunc.WriteMessage("Soft Bin of " + bin1.Name + " and " + bin2.Name + " are same, but at least one of other information is different.\n" +
                    bin1.Name + " would not generate BinTable...");
            }
            return false;
        }

        private static bool IsInsert(BinTableRow targetBin, string insertFlag)
        {
            IEnumerable<string> binItems = targetBin.ItemList.Split(',').Select(x => x.Trim().ToUpper());

            if (binItems.Any(x => x.Equals(insertFlag.ToUpper())))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private static void AddOrInsert(BinTableSheet binTable, BinTableRow bin)
        {
            BinTableRow targetBin = binTable.Rows.FirstOrDefault(x => x.Name.Equals(bin.Name, StringComparison.OrdinalIgnoreCase) && !x.IsBackup);
            if (targetBin != null)
            {
                targetBin.Op = "OR";
                foreach (string flag in bin.ItemList.Split(',').Select(x => x.Trim()))
                {
                    if (IsInsert(targetBin, flag))
                    {
                        targetBin.ItemList = targetBin.ItemList += "," + flag;
                        targetBin.Items.Add("T");
                    }
                }
            }
            else
            {
                binTable.AddRow(bin);
            }
        }
    }
}
