using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business.CopyXml
{
    public class NwireAp
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
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                File.Copy(fileInfo.FullName, Path.Combine(targetDir, fileInfo.Name));
            }
        }

        private bool IsTargetFile(FileInfo file)
        {

            if (file.Exists)
            {
                if (Regex.IsMatch(file.Name, @"FreeRunClk_TDR_TRUE_32Clk_8Idle\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }
                if (Regex.IsMatch(file.Name, @"FreeRunClk_differential\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }
                if (Regex.IsMatch(file.Name, @"UART_x3_RX\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }
                if (Regex.IsMatch(file.Name, @"UART_x3_TX\.xml", RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
