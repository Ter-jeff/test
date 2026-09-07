using AutogenCommandLine.CommandLineOptions;

using CommonLib.Extension;

using MyCommandLineLib;

namespace AutogenCommandLine.Tools.CheckScript
{
    public class CheckScriptCommandLineApp : CommandLineApplicationBase
    {
        public CheckScriptCommandLineApp()
        {
            ToolName = "CheckScript";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var autogenOptions = (CheckScriptOptions)options;

            if (string.IsNullOrEmpty(autogenOptions.Datalog))
            {
                string message = $"{autogenOptions.Datalog} is empty!";
                throw new CommandLineException(ToolName, message);
            }

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var checkScriptOptions = (CheckScriptOptions)options;
            if (options == null)
            {
                return this;
            }

            string inPutProg = checkScriptOptions.Prog?.Trim() ?? string.Empty;
            string outPutPath = checkScriptOptions.OutputFolder?.Trim() ?? string.Empty;
            CmdType cmdType = CheckType(checkScriptOptions.Type);
            if (cmdType.BinCut && !string.IsNullOrEmpty(checkScriptOptions.Datalog) && !string.IsNullOrEmpty(inPutProg) && !string.IsNullOrEmpty(outPutPath))
            {
                new BinCut(checkScriptOptions.Datalog.Trim(), inPutProg, outPutPath).RunBinCut();
            }
            if (cmdType.Efuse)
            {
                string inPutBdf = checkScriptOptions.Bdf?.Trim() ?? string.Empty;
                string inPutConf = checkScriptOptions.Config?.Trim() ?? string.Empty;
                _ = new EFuse(checkScriptOptions.Datalog.Trim(), inPutBdf, inPutConf, "", inPutProg, outPutPath);
                EFuse.RunEfuse();
            }
            return this;
        }

        private static CmdType CheckType(string types)
        {
            var cmdType = new CmdType();
            if (types.ContainsIgnoreCase("VALIDATION"))
            {
                cmdType.Validation = true;
            }

            if (types.ContainsIgnoreCase("EFUSE"))
            {
                cmdType.Efuse = true;
            }

            if (types.ContainsIgnoreCase("BinCut"))
            {
                cmdType.BinCut = true;
            }

            if (types.ContainsIgnoreCase("ALL"))
            {
                cmdType.BinCut = true;
                cmdType.Validation = true;
                cmdType.Efuse = true;
            }
            return cmdType;
        }
    }
}
