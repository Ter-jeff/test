using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business.CopyXml
{
    [ExcludeFromCodeCoverage]
    public class NwireLcd
    {
        public void GenerateFlow(string directory)
        {
            string targetDir = FolderStructure.DirXmlFiles;
            var srcDirectoryInfo = new DirectoryInfo(directory);
            List<FileInfo> filelst = srcDirectoryInfo.GetFiles("*", SearchOption.TopDirectoryOnly).ToList().FindAll(IsTargetFile);
            foreach (FileInfo fileInfo in filelst)
            {
                if (File.Exists(Path.Combine(targetDir, fileInfo.Name)))
                {
                    File.Delete(Path.Combine(targetDir, fileInfo.Name));
                }
                File.Copy(fileInfo.FullName, Path.Combine(targetDir, fileInfo.Name));
            }
        }

        private bool IsTargetFile(FileInfo file)
        {
            if (file.Exists)
            {
                if (Regex.IsMatch(file.Name, @"FreeRunClk_differential\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"FreeRunClk_TDR_TRUE_32Clk_8Idle\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"nWire_JTAG\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"nWire_JTAG_BitField\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"nWire_SPMI\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"nWire_SPMI_BitField\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"SPI\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"AV_JTAG\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"FRC\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                if (Regex.IsMatch(file.Name, @"JTAG_AHB_RW\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
