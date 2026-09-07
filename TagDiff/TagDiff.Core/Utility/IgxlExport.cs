using System;
using System.IO;
using System.IO.Compression;

using CommonLib.Extension;

namespace TagDiff.Core.Utility
{
    public static class IgxlExport
    {
        public static void ExportFiles(string testProgramPath, string exportPath)
        {
            if (string.IsNullOrWhiteSpace(testProgramPath))
            {
                throw new ArgumentException("testProgramPath is null or empty.", nameof(testProgramPath));
            }

            string folder = Path.Combine(Path.GetDirectoryName(testProgramPath) ?? string.Empty, exportPath);

            // Ensure folder is an absolute path for safety checks below
            string folderFullPath = Path.GetFullPath(folder);
            if (Directory.Exists(folderFullPath))
            {
                try
                {
                    Directory.Delete(folderFullPath, true);
                }
                catch (Exception ex)
                {
                    throw new IOException($"Failed to remove existing export folder '{folderFullPath}'.", ex);
                }
            }
            Directory.CreateDirectory(folderFullPath);

            ZipFile.ExtractToDirectory(testProgramPath, folder);

            foreach (string filePath in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
            {
                string? directory = Path.GetDirectoryName(filePath);
                string fileName = Path.GetFileName(filePath);
                string newFileName = fileName.Replace("%20", " ");
                if (directory != null)
                {
                    string newPath = Path.Combine(directory, newFileName);
                    if (!filePath.EqualsIgnoreCase(newPath))
                    {
                        File.Move(filePath, newPath);
                    }
                }
            }
        }
    }
}
