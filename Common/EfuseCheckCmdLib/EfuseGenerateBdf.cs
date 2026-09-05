using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.InputManager.Data;
using Automation.Static;

using EfuseCheckCmdLib.EFuse;

using TestPlanLib.Efuse.Input;

namespace EfuseCheckCmdLib
{
    public partial class EfuseGenerateBdf(EFuseInputData eFuseInputData, bool isFromCheckScript = false) : EfuseGenerate(eFuseInputData)
    {
        [GeneratedRegex(@"Follow Tracker sheet : if SVM disable fuse ""(?<svmValue>\S+)""", RegexOptions.IgnoreCase)]
        private static partial Regex SvmDisableKeyRegex();

        [GeneratedRegex(@"Follow Tracker sheet : if SVM enable fuse ""(?<svmValue>\S+)""", RegexOptions.IgnoreCase)]
        private static partial Regex SvmEnableKeyRegex();

        public readonly bool IsFromCheckScript = isFromCheckScript;

        public override void WorkFlow()
        {
            //Filter out SLT job and bira bank
            List<BitDefTable> bitDefTables = new EFusePreProcess().FilterBankForBitDefTableByArraySizeSheet([.. EFuseInputData.EfuseBitDefTables.Select(x => x.Copy())], EFuseInputData.EfuseArraySizeSheet);

            //Algorithm cond
            FinalBitDefTables = AddConfigRowToBitDefTable(bitDefTables);

            //Add dummy fuse for bank_mon
            FinalBitDefTables = AddDummyFuseForBankMon(FinalBitDefTables);

            //Generate EFUSE_BitDef_Table
            WriteEFuseBitDef(IsFromCheckScript);

            if (!IsFromCheckScript)
            {
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, EFuseConst.BitDefTableFileName);
            }

            //Generate Config_table and Config_table_SVM
            bool hasSvmKey = CheckIfSvmKeyWord();
            List<BitDefTable> configTablesInBitDef = GetConfigRowInBitDefTables();
            CheckIfHasBkmProcess();
            WriteConfigTable(configTablesInBitDef, hasSvmKey, IsFromCheckScript);

            //Generate BKM
            if (!IsFromCheckScript)
            {
                GenerateBkmInfoTable();
            }
        }

        private void WriteEFuseBitDef(bool isFromCheckScript = false)
        {
            string fileNameXls;
            string fileNameTxt;
            if (isFromCheckScript)
            {
                var file = new FileInfo(EfuseAlgorithmCheck.TestplanFile);
                if (file.Exists && file.DirectoryName != null)
                {
                    fileNameXls = Path.Combine(file.DirectoryName, EFuseConst.BitDefTableFileName + ".xlsx");
                    fileNameTxt = Path.Combine(file.DirectoryName, EFuseConst.BitDefTableFileName + ".txt");
                }
                else
                {
                    throw new Exception("Can not get testplan path");
                }
            }
            else
            {
                fileNameXls = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.BitDefTableFileName + ".xlsx");
                fileNameTxt = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.BitDefTableFileName + ".txt");
            }

            WriteEFuseBitDef(fileNameXls, fileNameTxt);
        }

        private void WriteConfigTable(List<BitDefTable> bitDefTables, bool hasSvmKey, bool isFromCheckScript = false)
        {
            string xlsPath;
            string txtPath;
            if (isFromCheckScript)
            {
                var file = new FileInfo(EfuseAlgorithmCheck.TestplanFile);
                if (file.Exists && file.DirectoryName != null)
                {
                    xlsPath = Path.Combine(file.DirectoryName, EFuseConst.ConfigTableFileName + ".xlsx");
                    txtPath = Path.Combine(file.DirectoryName, EFuseConst.ConfigTableFileName + ".txt");
                }
                else
                {
                    throw new Exception("Can not get testplan path");
                }
            }
            else
            {
                xlsPath = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableFileName + ".xlsx");
                txtPath = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableFileName + ".txt");
            }

            if (File.Exists(txtPath))
            {
                File.Delete(txtPath);
            }

            //print Config_table
            Regex svmDisableKey = SvmDisableKeyRegex();
            WriteConfigRowData(xlsPath, txtPath, svmDisableKey, hasSvmKey, bitDefTables);

            //print Config_table_SVM
            if (hasSvmKey)
            {
                xlsPath = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableSvmFileName + ".xlsx");
                txtPath = Path.Combine(FolderStructure.DirCommonSheets, EFuseConst.ConfigTableSvmFileName + ".txt");

                if (File.Exists(txtPath))
                {
                    File.Delete(txtPath);
                }

                Regex svmEnableKey = SvmEnableKeyRegex();
                WriteConfigRowData(xlsPath, txtPath, svmEnableKey, true, bitDefTables);
            }
        }
    }
}
