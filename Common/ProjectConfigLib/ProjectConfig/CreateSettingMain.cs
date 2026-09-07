using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using LogLib.Utility;

namespace ProjectConfigLib.ProjectConfig
{
    public partial class CreateSettingMain
    {
        [GeneratedRegex("_Default_RF$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("_Default_LCD$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("Cayman", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex("_Default$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex("Cayman", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();

        private string _existCopyFilename = "";

        public void CreateSettings(string folderName, string projectName, bool isAlarm, string type = "", bool overwrite = false)
        {
            var theFolder = new DirectoryInfo(folderName);
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();

            foreach (DirectoryInfo folder in dirInfo)
            {
                IOrderedEnumerable<FileInfo> files = folder.GetFiles().OrderByDescending(p => p.FullName);
                foreach (FileInfo file in files)
                {
                    if (type.EqualsIgnoreCase("RF"))
                    {
                        if (MyRegex().IsMatch(Path.GetFileNameWithoutExtension(file.FullName)))
                        {
                            CopyDefaultFileWithProject(file, projectName, "Default_RF", overwrite);
                            continue;
                        }
                    }
                    if (type.EqualsIgnoreCase("LCD"))
                    {
                        if (MyRegex1().IsMatch(Path.GetFileNameWithoutExtension(file.FullName)))
                        {
                            CopyDefaultFileWithProject(file, projectName, "Default_LCD", overwrite);
                            continue;
                        }
                    }
                    if (MyRegex3().IsMatch(Path.GetFileNameWithoutExtension(file.FullName)))
                    {
                        CopyDefaultFileWithProject(file, projectName, "Default", overwrite);
                    }
                }

                DirectoryInfo[] dirs = folder.GetDirectories();
                foreach (DirectoryInfo dir in dirs)
                {
                    if (MyRegex4().IsMatch(dir.FullName))
                    {
                        string newFolder = MyRegex2().Replace(dir.FullName, projectName);
                        if (!Directory.Exists(newFolder))
                        {
                            Directory.CreateDirectory(newFolder);
                        }
                        else
                        {
                            if (isAlarm)
                            {
                                string message = "Settings folder " + newFolder + "  is already exist!";
                                ErrorMessageBox.Show(message);
                            }
                        }
                    }
                }
            }
        }

        private void CopyDefaultFileWithProject(FileInfo fileInfo, string projectName, string defaultPattern, bool overwrite = false)
        {
            string newFile = Regex.Replace(fileInfo.FullName, defaultPattern, projectName, RegexOptions.IgnoreCase);
            if (_existCopyFilename.EqualsIgnoreCase(newFile))
            {
                return;
            }

            if (!File.Exists(newFile))
            {
                File.Copy(fileInfo.FullName, newFile);
                _existCopyFilename = newFile;
            }
            else
            {
                if (overwrite)
                {
                    File.Copy(fileInfo.FullName, newFile);
                    _existCopyFilename = newFile;
                }
            }
        }
    }
}
