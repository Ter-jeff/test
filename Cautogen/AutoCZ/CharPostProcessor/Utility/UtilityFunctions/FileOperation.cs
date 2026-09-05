using System.IO;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions
{
    public static class FileOperation
    {
        public static void CheckFolderExist(string folder)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }
    }
}
