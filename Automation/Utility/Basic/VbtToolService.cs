using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

namespace Automation.Utility.Basic
{
    public static class VbtToolService
    {
        public static void ModifyCommonBas(List<FileInfo> fileInfos)
        {
            if (!fileInfos.Any())
            {
                return;
            }

            FileInfo targetBas = fileInfos.Find(p => p.Name.Equals("LIB_Common_GlobalConstant.bas", StringComparison.OrdinalIgnoreCase));
            List<string> lines = File.ReadAllLines(targetBas.FullName).ToList();
            int dcviIndex = lines.FindIndex(x => x.Equals(SearchContent(lines, new List<string> { "AllDCVIPinlist", "Const" })));
            int powerIndex = lines.FindIndex(p => p.Equals(SearchContent(lines, new List<string> { "AllPowerPinlist", "Const" })));
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                if (!lines[dcviIndex].Split(' ').Last().Equals("\"DCVI_Power\"", StringComparison.OrdinalIgnoreCase))
                {
                    lines[dcviIndex] = "Public Const AllDCVIPinlist = \"DCVI_Power\"";
                }
                if (!lines[powerIndex].Split(' ').Last().Equals("\"DCVS_Power\"", StringComparison.OrdinalIgnoreCase))
                {
                    lines[powerIndex] = "Public Const AllPowerPinlist = \"DCVS_Power\"";
                }
            }
            else
            {
                if (!lines[dcviIndex].Split(' ').Last().Equals("\"All_DCVI\"", StringComparison.OrdinalIgnoreCase))
                {
                    lines[dcviIndex] = "Public Const AllDCVIPinlist = \"All_DCVI\"";
                }
                if (!lines[powerIndex].Split(' ').Last().Equals("\"All_Power\"", StringComparison.OrdinalIgnoreCase))
                {
                    lines[powerIndex] = "Public Const AllDCVIPinlist = \"All_Power\"";
                }
            }
            File.WriteAllLines(targetBas.FullName, lines);
        }

        public static List<FileInfo> GetVbFiles(string dirLib, string dirCodingLib)
        {
            var mFileList = new List<FileInfo>();
            if (LocalSpecs.Options.Device == EnumDevice.AP)
            {
                mFileList = GetLibList(dirLib);
            }
            else if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                mFileList = GetLibList(dirLib);
                var cordingLib = new DirectoryInfo(dirCodingLib);
                mFileList.AddRange(GetVbtFilePath(cordingLib));
            }
            else if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                mFileList = GetLibListNonAp(dirLib);
                mFileList.AddRange(GetLibList(Path.Combine(dirLib, "Wireless")));
                ModifyCommonBas(mFileList);
            }

            return mFileList;
        }

        private static string SearchContent(List<string> content, List<string> patterns)
        {
            bool hasSearch = true;
            foreach (string perLine in content)
            {
                foreach (string pattern in patterns)
                {
                    if (!perLine.ContainsIgnoreCase(pattern))
                    {
                        hasSearch = false;
                        break;
                    }
                    hasSearch = true;
                }
                if (hasSearch)
                {
                    return perLine;
                }
            }
            return "";
        }

        public static List<FileInfo> GetLibList(string folder)
        {
            if (!Directory.Exists(folder))
            {
                return new List<FileInfo>();
            }

            var dir = new DirectoryInfo(folder);
            List<FileInfo> mFileList = dir.GetFiles("*.bas", SearchOption.TopDirectoryOnly).ToList();
            mFileList.AddRange(dir.GetFiles("*.cls", SearchOption.TopDirectoryOnly).ToList());
            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                mFileList.AddRange(subDir.GetFiles("*.bas", SearchOption.TopDirectoryOnly));
                mFileList.AddRange(subDir.GetFiles("*.cls", SearchOption.TopDirectoryOnly));
            }
            return mFileList;
        }

        public static List<FileInfo> GetLibListNonAp(string folder)
        {
            var unusedModule = new List<string>
            {
                "eFuse",
                "PFA_SONE",
                "SPIROM",
                "TMPS_ADCTrim_FreqSync_Det",
                "VDDBinning"
            };
            var dir = new DirectoryInfo(folder);
            List<FileInfo> mFileList = dir.GetFiles("*.bas", SearchOption.TopDirectoryOnly).ToList();
            mFileList.AddRange(dir.GetFiles("*.cls", SearchOption.TopDirectoryOnly).ToList());

            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                if (!unusedModule.Exists(s => s.Equals(subDir.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    mFileList.AddRange(subDir.GetFiles("*.bas", SearchOption.TopDirectoryOnly).Where(p => !Regex.IsMatch(p.Name, "_AP.")).ToList());
                    mFileList.AddRange(subDir.GetFiles("*.cls", SearchOption.TopDirectoryOnly).Where(p => !Regex.IsMatch(p.Name, "_AP.")).ToList());
                }
            }
            return mFileList;
        }

        public static List<FileInfo> GetVbtFilePath(DirectoryInfo dir)
        {
            List<FileInfo> mFileList = dir.GetFiles("*.bas", SearchOption.TopDirectoryOnly).ToList();
            mFileList.AddRange(dir.GetFiles("*.cls", SearchOption.TopDirectoryOnly).ToList());
            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                mFileList.AddRange(subDir.GetFiles("*.bas", SearchOption.TopDirectoryOnly));
                mFileList.AddRange(subDir.GetFiles("*.cls", SearchOption.TopDirectoryOnly));
            }
            return mFileList;
        }
    }
}
