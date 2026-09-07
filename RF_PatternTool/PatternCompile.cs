using System.Diagnostics;

namespace PatternCompile
{
    public class PatternConv
    {
        public static void ZipFile(string fileName)
        {
            Process process = new Process();
            ProcessStartInfo info = new Process().StartInfo;
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.FileName = "gzip.exe";
            info.Arguments = "\"" + fileName + "\" -f";
            process.StartInfo = info;
            process.Start();
            process.WaitForExit();
        }

        public static void CompileToPat(string fileName, string pinmap)
        {
            Process process = new Process();
            ProcessStartInfo info = new Process().StartInfo;
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.FileName = "apc.exe";
            info.Arguments = string.Format("\"{0}\" -pinmap_workbook \"{1}\" -comments -digital_inst {2} -import_all_undefineds", fileName + ".gz", pinmap, "HSDMQ");
            process.StartInfo = info;
            process.Start();
            process.WaitForExit();
        }

        public static void GetPatternInfo(string folder, string filename)
        {

            string subfolder = Path.GetFileNameWithoutExtension(filename);
            string patfolder = Path.Combine(folder, subfolder);

            if (Directory.Exists(patfolder))
            {
                Directory.Delete(patfolder, true);
            }

            Directory.CreateDirectory(patfolder);

            File.Copy(filename + ".gz", Path.Combine(folder, subfolder, Path.GetFileName(filename) + ".gz"));

            Process process = new Process();
            ProcessStartInfo info = new Process().StartInfo;
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.CreateNoWindow = true;
            info.FileName = Path.Combine(Path.GetDirectoryName(AppContext.BaseDirectory), "run.exe");
            info.UseShellExecute = true;
            info.Arguments = $"-d \"{patfolder}\" -w";
            process.StartInfo = info;
            process.Start();
            process.WaitForExit();

        }

        public static List<string> GetInfoFromPatternFolder(string folder)
        {
            var result = new List<string>();
            #region Read Reference info
            string[] files = Directory.GetFiles(folder, "*.txt");
            foreach (string file in files)
            {
                var rd = new StreamReader(file);
                string line;
                while ((line = rd.ReadLine()) != null)
                {
                    result.Add(line);
                }
                rd.Close();
            }

            #endregion

            return result;
        }
    }
}
