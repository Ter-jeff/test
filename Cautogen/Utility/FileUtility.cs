using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Cautogen.Utility
{
    public static class FileUtility
    {
        public static void CleanDir(string dirPath)
        {
            if (Directory.Exists(dirPath))
            {
                var dirInfo = new DirectoryInfo(dirPath);
                foreach (FileInfo file in dirInfo.GetFiles())
                {
                    file.Delete();
                }

                foreach (DirectoryInfo dir in dirInfo.GetDirectories())
                {
                    dir.Delete(true);
                }
            }
            else
            {
                Directory.CreateDirectory(dirPath);
            }
        }

        public static bool CompareFileList(List<string> fileListExpect, List<string> fileListResult, string logPath = "")
        {
            var expectDic = fileListExpect.ToDictionary(Path.GetFileName, p => p);
            var resultDic = fileListResult.ToDictionary(Path.GetFileName, p => p);

            var newItems = new List<string>();
            var lackItems = new List<string>();
            var md5DiffItems = new List<Tuple<string, string>>();

            bool result = true;

            foreach (string file in expectDic.Keys.Where(file => !resultDic.ContainsKey(file)))
            {
                lackItems.Add(file);
                result = false;
            }

            foreach (string file in resultDic.Keys.Where(file => !expectDic.ContainsKey(file)))
            {
                newItems.Add(file);
                result = false;
            }

            foreach (string fileName in
                from fileName in resultDic.Keys
                where !Regex.IsMatch(fileName, "manifest", RegexOptions.IgnoreCase)
                where !newItems.Contains(fileName)
                where !_AreFilesEqual(expectDic[fileName], resultDic[fileName])
                where !string.Equals(fileName, "tl_WorkBookProperties_.txt")  // ignore workbook properities
                where !string.Equals(fileName, "IGLinkManifest.txt")
                select fileName)
            {
                md5DiffItems.Add(new Tuple<string, string>(expectDic[fileName], resultDic[fileName]));
                result = false;
            }

            if (logPath == "")
            {
                return result;
            }

            using (var file = new StreamWriter(logPath, true))
            {
                foreach (Tuple<string, string> diffpair in md5DiffItems)
                {
                    file.WriteLine(@"Different Sheet: '{0}' \n", diffpair);
                }

                foreach (string item in lackItems)
                {
                    file.WriteLine(@"Lack Sheet in Result: '{0}' \n", item);
                }

                foreach (string item in newItems)
                {
                    file.WriteLine(@"Extra Sheet in Result: '{0}' \n", item);
                }
            }
            return result;
        }

        public static string GetFirstFile(string folderPath, string searchStr)
        {
            return Directory.GetFiles(folderPath).FirstOrDefault(filePath => filePath.Contains(searchStr));
        }

        public static List<string> GetFiles(string folderPath, string filter = "*")
        {
            var basFilePaths = new List<string>();
            if (!Directory.Exists(folderPath))
            {
                return basFilePaths;
            }

            // get filepath in the current foulder
            basFilePaths.AddRange(Directory.GetFiles(folderPath, filter));

            // get file path in sub-folers
            foreach (string d in Directory.GetDirectories(folderPath))
            {
                basFilePaths.AddRange(GetFiles(d, filter));
            }
            return basFilePaths;
        }

        public static string GetFileContains(string folderPath, string containString)
        {
            return Directory.GetFiles(folderPath).FirstOrDefault(filePath => filePath.Contains(containString));
        }

        private static string _GetFile(string filename)
        {
            string line = "";
            StringBuilder totalline = new StringBuilder("");

            var filestream = new FileStream(filename, FileMode.Open, FileAccess.Read);
            using (var textReader = new StreamReader(filestream))
            {
                while ((line = textReader.ReadLine()) != null)
                {
                    totalline.Append(line);
                    totalline.Append("\n");
                }
            }
            return totalline.ToString();
        }
        private static bool _AreFilesEqual(string first, string second)
        {
            return _GetFile(first).Equals(_GetFile(second));
        }
    }
}
