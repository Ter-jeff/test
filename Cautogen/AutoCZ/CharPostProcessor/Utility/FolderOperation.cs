using System.IO;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility
{
    public static class FolderOperation
    {

        public static void CopyFilesRecursively(string sourcePath, string targetPath)
        {
            //Now Create all of the directories
            foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
            }

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
            }
        }

        public static string GetParentDirectory(string path, int levels)
        {
            for (int i = 0; i < levels; i++)
            {
                path = Path.GetDirectoryName(path);
                if (path == null)
                {
                    return null;
                }
            }
            return path;
        }
    }
}
