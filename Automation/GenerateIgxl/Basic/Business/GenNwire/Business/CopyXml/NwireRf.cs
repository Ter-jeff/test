using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business.CopyXml
{
    public class NwireRf
    {
        public void GenerateFlow(string directory)
        {
            string targetDir = FolderStructure.DirXmlFiles;
            DirectoryInfo srcDirectoryInfo = new DirectoryInfo(directory);
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
                if (Regex.IsMatch(file.Name, @"FRC\.pa", RegexOptions.IgnoreCase))
                {
                    return true;
                }
                if (Regex.IsMatch(file.Name, @"FRCRef\.pa", RegexOptions.IgnoreCase))
                {
                    return true;
                }
                if (Regex.IsMatch(file.Name, @"FRCRef_differential\.pa", RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
