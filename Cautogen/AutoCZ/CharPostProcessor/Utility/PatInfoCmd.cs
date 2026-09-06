using System;
using System.Diagnostics;
using System.IO;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility
{
    public class PatInfoCmd
    {
        // Get information from #.Pat
        public bool ConvertByArgs(string sourcePat, ref string output, string strArg)
        {
            string args = $" {strArg} \"{sourcePat}\"";
            try
            {
                var p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.FileName = Path.Combine(Environment.GetEnvironmentVariable("igxlroot"), "bin", "patinfo.exe");
                p.StartInfo.Arguments = args;
                p.Start();

                output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                if (p.ExitCode == 0)
                {
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
