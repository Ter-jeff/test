using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Enums;
using CommonLib.Extension;
using CommonLib.Static;

using LogLib.Static;
using LogLib.Utility;

using Moq;

using MyCommandLineLib.Enums;
using MyCommandLineLib.Extensions;

using OfficeOpenXml.Packaging.Ionic.Zlib;

namespace MyCommandLineLib
{
    public abstract class CommandLineApplicationBase : ICommandLineApplication
    {
        public static readonly DateTime DefaultDate = new(2022, 01, 01, 00, 00, 00, DateTimeKind.Utc);
        public bool UseStateMachine { get; set; } = true;
        public string ToolName { get; set; } = "";

        public virtual ICommandLineOptions Start(ICommandLineOptions commandLineOptions)
        {
            Response.Report("", EnumMessageLevel.General, 0);
            Response.Report("****" + (ToolName + " Start").PadCenter(40) + "****", EnumMessageLevel.General, 0);
            return commandLineOptions;
        }

        public virtual ICommandLineOptions End(ICommandLineOptions commandLineOptions)
        {
            Response.Report("****" + (ToolName + " End").PadCenter(40) + "****", EnumMessageLevel.General, 0);
            Response.Report("", EnumMessageLevel.General, 0);
            return commandLineOptions;
        }

        public virtual ICommandLineOptions PreAction(ICommandLineOptions commandLineOptions)
        {
            string logName = "TSMC" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".log";
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (!File.Exists(folder))
            {
                folder = Directory.GetCurrentDirectory();
            }
            string file = Path.Combine(folder, "Teradyne", logName);
            LogHelper.SetNLog(file);

            return commandLineOptions;
        }

        public abstract ICommandLineOptions ValidateInput(ICommandLineOptions commandLineOptions);

        public abstract ICommandLineApplication Execute(ICommandLineOptions commandLineOptions);

        public virtual ICommandLineApplication PostAction(ICommandLineOptions commandLineOptions)
        {
            if (!string.IsNullOrEmpty(commandLineOptions.GetOutputFolder()))
            {
                string logFileName = ToolName + ".log";
                LogHelper.Copy(Path.Combine(commandLineOptions.GetOutputFolder(), logFileName));
                LogHelper.DeleteLogFile();
            }
            return this;
        }

        protected bool CheckStdFilePath(string std, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(std))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing std file path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!std.EndsWithIgnoreCase(".std"))
            {
                string message = string.Format("The input std file {0} is not a std file (.std) at argument " + symbol + ", -?\" to get help.", std);
                throw new CommandLineException(ToolName, message);
            }
            if (!File.Exists(std))
            {
                string message = $"The csv file {std} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckCsvFilePath(string csv, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(csv))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing cvs file path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!csv.EndsWithIgnoreCase(".csv"))
            {
                string message = string.Format("The input csv file {0} is not a csv file (.csv) at argument " + symbol + ", -?\" to get help.", csv);
                throw new CommandLineException(ToolName, message);
            }
            List<string> csvFiles = [.. csv.Split(',')];
            foreach (string csvFile in csvFiles)
            {
                if (!File.Exists(csvFile))
                {
                    string message = $"The csv file {csv} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }
            return true;
        }

        protected bool CheckYamlFilePath(string yaml, bool required = true)
        {
            if (string.IsNullOrEmpty(yaml))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing yaml file path at argument -y, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!yaml.EndsWithIgnoreCase(".yaml"))
            {
                string message =
                    $"The input yaml file {yaml} is not a yaml file (.yaml) at argument -y, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!File.Exists(yaml))
            {
                string message = $"The yaml file {yaml} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckOtpFilesPath(string otpFiles, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(otpFiles))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing otp Files path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            List<string> otps = [.. otpFiles.Split(',')];
            foreach (string otp in otps)
            {
                if (!otp.EndsWithIgnoreCase(".otp"))
                {
                    string message =
                        string.Format("The input otp file {0} is not a yaml file (.otp) at argument " + symbol + ", -?\" to get help.", otp);
                    throw new CommandLineException(ToolName, message);
                }

                if (!File.Exists(otp))
                {
                    string message = $"The otp file {otp} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }

            return true;
        }

        protected bool CheckTestProgramFileFolderPath(string tpPath, bool required = true)
        {
            if (string.IsNullOrEmpty(tpPath))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing igxl test program file path at argument -t, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!File.Exists(tpPath) && !Directory.Exists(tpPath))
            {
                string message = $"Test program file {tpPath} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckTestProgramFilePath(string tpPath, bool required)
        {
            if (string.IsNullOrEmpty(tpPath))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing igxl test program file path at argument -t, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!tpPath.EndsWithIgnoreCase(".igxl"))
            {
                string message =
                    $"The input test program file {tpPath} is not an igxl test program file (.igxl) at argument -t, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!File.Exists(tpPath))
            {
                string message = $"Test program file {tpPath} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckTemplateFilesPath(string templates, bool required = true)
        {
            if (string.IsNullOrEmpty(templates))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing template files path at argument -t, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            List<string> templateFiles = [.. templates.Split(',')];
            foreach (string template in templateFiles)
            {
                if (!template.EndsWithIgnoreCase(".tmp"))
                {
                    string message =
                        $"The input template file {template} is not an template file (.tmp) at argument -t, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
                if (!File.Exists(template))
                {
                    string message = $"Test template file {template} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }
            return true;
        }

        protected bool CheckOutputFolderPath(string outputDir, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(outputDir))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing folder path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            return true;
        }

        protected bool CheckFolderPathOrTxt(string folder, string symbol, bool required = true)
        {
            if (!required)
            {
                return true;
            }

            if (string.IsNullOrEmpty(folder))
            {
                string message = "Missing file or path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            bool isFile = folder.EndsWithIgnoreCase(".txt");
            if (isFile)
            {
                if (!File.Exists(folder))
                {
                    string message = $"The file {folder} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }
            else
            {
                if (!Directory.Exists(folder))
                {
                    string message = $"Test folder {folder} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }
            return true;
        }

        protected bool CheckFolderPath(string folder, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(folder))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing folder path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            if (!Directory.Exists(folder))
            {
                string message = $"Test folder {folder} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckJson(string file, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(file))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing file at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            if (!File.Exists(file))
            {
                string message = $"File {file} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            if (!file.EndsWithIgnoreCase(".json"))
            {
                string message = "The extension of file is not .json";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckExcelFilePath(string excels, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(excels))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing excel files path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            foreach (string excel in excels.Split(','))
            {
                if (!excel.EndsWithIgnoreCase(".xlsx") &&
                    !excel.EndsWithIgnoreCase(".xlsm"))
                {
                    string message = string.Format("The input excel file {0} is not an excel file (.xlsx or .xlsm) at argument " + symbol + ", -?\" to get help.", excel);
                    throw new CommandLineException(ToolName, message);
                }
                if (!File.Exists(excel))
                {
                    string message = $"Test excel file {excel} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }
            return true;
        }

        protected bool CheckBaseFilesPath(string basFiles, bool required = true)
        {
            if (string.IsNullOrEmpty(basFiles))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing bas files path at argument -b, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }

            foreach (string bas in basFiles.Split(',').ToList())
            {
                if (!bas.EndsWithIgnoreCase(".bas"))
                {
                    string message =
                        $"The input bas file {bas} is not an bas file (.bas) at argument -b, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
                if (!File.Exists(bas))
                {
                    string message = $"Test bas file {bas} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }
            return true;
        }

        protected bool CheckOutputFilePath(string outputPath)
        {
            if (string.IsNullOrEmpty(outputPath))
            {
                string message = "Missing output file path at argument -o, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            try
            {
                if (!Directory.Exists(outputPath))
                {
                    if (outputPath != null)
                    {
                        Directory.CreateDirectory(outputPath);
                    }
                }
            }
            catch (Exception e)
            {
                string message =
                    $"Create output directory error, maybe the output path {outputPath} is not valid: {e.StackTrace}";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckOutputFolder(string outputFolder)
        {
            if (string.IsNullOrEmpty(outputFolder))
            {
                string message = "Missing output folder at argument -o, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            try
            {
                if (!Directory.Exists(outputFolder))
                {
                    if (outputFolder != null)
                    {
                        Directory.CreateDirectory(outputFolder);
                    }
                }
            }
            catch (Exception e)
            {
                string message =
                    $"Create output directory error, maybe the output path {outputFolder} is not valid: {e.StackTrace}";
                throw new CommandLineException(ToolName, message);
            }

            return true;
        }

        protected bool CheckLogPath(string logPath)
        {
            if (string.IsNullOrEmpty(logPath))
            {
                return true;
            }

            try
            {
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
            }
            catch (Exception e)
            {
                string message =
                    $"Create log path directory error, maybe the log path {logPath} is not valid: {e.StackTrace}";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckBoolean(int function, string symbol)
        {
            if (!function.Equals(0) && !function.Equals(1))
            {
                string message = "The value at argument " + symbol + " is not valid, it must be 1 or 0, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckList(string text, List<string> list, string symbol)
        {
            if (!list.Exists(x => x.EqualsIgnoreCase(text)))
            {
                string message = "The value at argument " + symbol + " is not valid, it must be " + string.Join(",", list) + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }

        protected bool CheckTxtFilePath(string filePath, string symbol, bool required = true)
        {
            return CheckGenericFilePath(EnumFileExtension.Txt, filePath, symbol, required);
        }

        protected bool CheckIniFilePath(string filePath, string symbol, bool required = true)
        {
            return CheckGenericFilePath(EnumFileExtension.Ini, filePath, symbol, required);
        }

        protected bool CheckXmlFilePath(string filePath, string symbol, bool required = true)
        {
            return CheckGenericFilePath(EnumFileExtension.Xml, filePath, symbol, required);
        }

        protected bool CheckLogFilePath(string filePath, string symbol, bool required = true)
        {
            return CheckGenericFilePath(EnumFileExtension.Log, filePath, symbol, required);
        }

        public static void Mock()
        {
            TimeMock();

            AssemblyMock();
        }

        private static void TimeMock()
        {
            TimeContext.ResetToDefault();
            var timeMock = new Mock<TimeProvider>();
            timeMock.Setup(tp => tp.GetUtcNow()).Returns(DefaultDate);
            timeMock.SetupGet(tp => tp.LocalTimeZone).Returns(TimeZoneInfo.Utc);
            TimeContext.Current = timeMock.Object;
        }

        private static void AssemblyMock()
        {
            AssemblyProvider.Current = MockAssemblyProvider.Instance;
        }

        protected static Stream? GetStream(string log)
        {
            string ext = Path.GetExtension(log);
            if (ext.EqualsIgnoreCase(".gz"))
            {
                FileStream compressedFileStream = File.OpenRead(log);
                var decompressionStream = new GZipStream(compressedFileStream, CompressionMode.Decompress);
                return decompressionStream;
            }
            else if (ext.EqualsIgnoreCase(".txt"))
            {
                FileStream inputStream = File.OpenRead(log);
                return inputStream;
            }
            Response.Report(log + " only support *.txt or *.gz ", EnumMessageLevel.Error, 0);
            return null;
        }

        private bool CheckGenericFilePath(EnumFileExtension enumFileExtension, string filePath, string symbol, bool required = true)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                if (!required)
                {
                    return true;
                }

                string message = "Missing file path at argument " + symbol + ", -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            if (!filePath.EndsWithIgnoreCase(enumFileExtension.GetExtension()))
            {
                string message = string.Format("The input file {0} is not .txt at argument " + symbol + ", -?\" to get help.", filePath);
                throw new CommandLineException(ToolName, message);
            }
            if (!File.Exists(filePath))
            {
                string message = $"The file {filePath} not found, -?\" to get help.";
                throw new CommandLineException(ToolName, message);
            }
            return true;
        }
    }
}
