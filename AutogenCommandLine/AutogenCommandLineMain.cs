using System;
using System.Globalization;
using System.Text;

using Automation;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;

using MyCommandLineLib;

namespace AutogenCommandLine
{
    public class AutogenCommandLineMain
    {
        public static int Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            if (!EpplusExtensions.IsNativeAutoFitAvailable())
            {
                Response.Report(
                    "libgdiplus/GDI+ not detected — Excel report column widths are estimated (SkiaSharp) and may differ slightly from native auto-fit.",
                    EnumMessageLevel.Warning);
            }

            try
            {
                if (!CommandLineManager.ParseArgList(args, out ICommandLineOptions iCommandLineOptions, out EntryOptions entryOptions))
                {
                    return 0;
                }

                if (iCommandLineOptions != null)
                {
                    CommandLineManager.Execute(iCommandLineOptions, entryOptions);
                    return GenerateIgxlMain.ReturnValue;
                }
            }
            catch (CommandLineException ex)
            {
                Console.Error.WriteLine(DateTime.Now.ToLocalTime().ToString(CultureInfo.InvariantCulture) + ": {0} error: {1}", ex.Caller, ex.Message);
                return 1;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(DateTime.Now.ToLocalTime().ToString(CultureInfo.InvariantCulture) + ": CommandLine error: {0}", ex.Message);
                return 1;
            }
            return 0;
        }
    }
}
