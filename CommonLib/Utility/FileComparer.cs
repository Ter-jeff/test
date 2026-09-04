using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace CommonLib.Utility
{
    public class DiffPair
    {
        public string ExpectedFile = "";
        public string OutputFile = "";

        public DiffPair(string output, string expected)
        {
            ExpectedFile = expected;
            OutputFile = output;
        }
    }

    public class FileComparer
    {
        private const int BytesToRead = sizeof(long);
        public List<string> NewItems = new List<string>();
        public List<string> LackItems = new List<string>();
        public List<DiffPair> Md5DiffItems = new List<DiffPair>();

        public static bool CompareFile(string fileName)
        {
            return false;
        }

        public static bool FilesAreEqual(FileInfo first, FileInfo second)
        {
            if (first.Length != second.Length)
            {
                return false;
            }

            int iterations = (int)Math.Ceiling((double)first.Length / BytesToRead);

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead())
            {
                byte[] one = new byte[BytesToRead];
                byte[] two = new byte[BytesToRead];

                for (int i = 0; i < iterations; i++)
                {
                    fs1.Read(one, 0, BytesToRead);
                    fs2.Read(two, 0, BytesToRead);

                    if (BitConverter.ToInt64(one, 0) != BitConverter.ToInt64(two, 0))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        public bool CompareFileList(List<string> fileListExpect, List<string> fileListResult)
        {
            NewItems = new List<string>();
            LackItems = new List<string>();
            Md5DiffItems = new List<DiffPair>();
            bool result = true;

            foreach (string file in fileListExpect)
            {
                string chkFile = file.Replace("\\Expected\\", "\\Output\\");
                if (!File.Exists(chkFile))
                {
                    LackItems.Add(file);
                    result = false;
                }
            }
            foreach (string file in fileListResult)
            {
                string chkFile = file.Replace("\\Output\\", "\\Expected\\");
                if (!File.Exists(chkFile))
                {
                    NewItems.Add(file);
                    result = false;
                }
            }

            foreach (string fileName in fileListResult)
            {
                string expectFileName = fileName.Replace("\\Output\\", "\\Expected\\");
                if (!NewItems.Contains(fileName))
                {
                    if (!FilesAreEqualUsingMd5(fileName, expectFileName))
                    {
                        var diffpair = new DiffPair(fileName, expectFileName);
                        Md5DiffItems.Add(diffpair);
                        result = false;
                    }
                }
            }
            return result;
        }

        public FileCompareResult CompareFileListNotInSameFolderStruct(List<string> fileListExpect, List<string> fileListResult)
        {
            var fileCompResult = new FileCompareResult();
            fileCompResult.Result = true;
            foreach (string file in fileListExpect)
            {
                string chkFile = fileListResult.Find(s => GetFileName(s).Equals(GetFileName(file), StringComparison.OrdinalIgnoreCase));
                if (!File.Exists(chkFile))
                {
                    fileCompResult.LackItems.Add(file);
                    fileCompResult.Result = false;
                }
            }
            foreach (string file in fileListResult)
            {
                string chkFile = fileListExpect.Find(s => GetFileName(s).Equals(GetFileName(file), StringComparison.OrdinalIgnoreCase));
                if (!File.Exists(chkFile))
                {
                    fileCompResult.NewItems.Add(file);
                    fileCompResult.Result = false;
                }
            }

            foreach (string fileName in fileListResult)
            {
                string expectFileName = fileListExpect.Find(s => GetFileName(s).Equals(GetFileName(fileName), StringComparison.OrdinalIgnoreCase));
                if (!NewItems.Contains(fileName))
                {
                    if (!FilesAreEqualUsingMd5(fileName, expectFileName))
                    {
                        var diffpair = new DiffPair(fileName, expectFileName);
                        fileCompResult.MD5diffItems.Add(diffpair);
                        fileCompResult.Result = false;
                    }
                }
            }
            return fileCompResult;
        }

        public bool CompareFile(string fileName, string expectFileName)
        {
            if (!File.Exists(fileName) || !File.Exists(expectFileName))
            {
                return false;
            }

            if (!FilesAreEqualUsingMd5(fileName, expectFileName))
            {
                var diffpair = new DiffPair(fileName, expectFileName);
                Md5DiffItems.Add(diffpair);
                return false;
            }
            return true;
        }

        public bool FilesAreEqualUsingMd5(string first, string second)
        {
            string md5ForFirst = GetFileMD5(first);
            string md5ForSecond = GetFileMD5(second);

            if (md5ForFirst.Equals(md5ForSecond))
            {
                return true;
            }
            return false;
        }

        private static string GetFileMD5(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = new StreamReader(stream))
                    {
                        string txt = "";
                        foreach (byte byte_ in md5.ComputeHash(stream))
                        {
                            txt = string.Format("{0}{1}", txt, byte_.ToString("x2"));
                        }

                        return txt;
                    }
                }
            }
        }

        private static string GetFileName(string fileFullName)
        {
            return new FileInfo(fileFullName).Name;
        }
    }

    public class FileCompareResult
    {
        public bool Result;
        public List<string> NewItems = new List<string>();
        public List<string> LackItems = new List<string>();
        public List<DiffPair> MD5diffItems = new List<DiffPair>();
    }
}
