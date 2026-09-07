using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class RemoveDuplicateInstance
    {
        public static void WorkFlow(string path)
        {
            string folder = Path.Combine(path, ConstData.TrunkFolder);
            if (!Directory.Exists(folder))
            {
                return;
            }

            if (LocalSpecs.FileStructure.Count == 0)
            {
                return;
            }

            List<FileInfo> fileList = new DirectoryInfo(folder).EnumerateFiles("*",
                new EnumerationOptions
                {
                    RecurseSubdirectories = true,
                    AttributesToSkip = FileAttributes.ReparsePoint
                }).ToList();

            var duplicateList = fileList.GroupBy(x => x.Name).Where(x => x.Count() > 1).ToList();

            foreach (IGrouping<string, FileInfo> duplicate in duplicateList)
            {
                var sortByTime = duplicate.OrderByDescending(x => x.LastWriteTime).ToList();
                var deleteList = sortByTime.Skip(1).ToList();
                deleteList.ForEach(x => File.Delete(x.FullName));
            }
        }
    }
}
