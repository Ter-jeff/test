using System;
using System.Collections.Generic;
using System.Diagnostics;

using AutogenCommandLine.CommandLineOptions;

using Automation;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport.Base;
using CommonLib.Extension;

using CommonReaderLib;

using LcdLib;

using LogLib.Static;

using MyCommandLineLib;
using MyCommandLineLib.Enums;

using OfficeOpenXml;

using RfLib;

namespace AutogenCommandLine.Tools.Autogen
{
    public class InputInfoCommandLineApp : AutogenCommandLineApp
    {
        public InputInfoCommandLineApp()
        {
            ToolName = "InputInfo";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var arguments = (AutogenOptions)options;

            UseStateMachine &= CheckCsvFilePath(arguments.InputInfo, "--inputinfo");

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            try
            {
                var arguments = (AutogenOptions)options;
                if (arguments.Mock == 1)
                {
                    Mock();
                }
                AllStatic.Clear();
                LocalSpecs.IsUnitTest = arguments.Mock == 1;

                string inputInfo = arguments.InputInfo;
                using var package = new ExcelPackage();
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("InputInfo");
                List<List<string>> lists = inputInfo.CsvConvertToLists();
                _ = ws.Cells[1, 1].PrintExcelRowByList(lists);
                InputInfoSheet sheet = new InputInfoReader(inputInfo).ReadSheet(ws);
                sheet.Check();
                if (sheet.GetErrors().Count == 0)
                {
                    var userInfo = new UserInfo(sheet);
                    userInfo.Initialize();

                    if (userInfo.Entry.Equals(nameof(EnumEntryType.Autogen), StringComparison.CurrentCultureIgnoreCase))
                    {
                        new GenerateIgxlMain().Run();
                    }
                    else if (userInfo.Entry.Equals(nameof(EnumEntryType.Validation), StringComparison.CurrentCultureIgnoreCase))
                    {
                        new ValidateTestPlanMain().Run();
                    }
                    else if (LocalSpecs.Options.Device == EnumDevice.RF)
                    {
                        GenerateIgxlMainRf.Run();
                    }
                    else if (LocalSpecs.Options.Device == EnumDevice.LCD)
                    {
                        new GenerateIgxlMainLcd().Run();
                    }

                    PostWorkFlow(LocalSpecs.IsUnitTest ? 1 : 0, userInfo.OutputFolder);
                }
                else
                {
                    foreach (Error error in sheet.GetErrors())
                    {
                        Response.Report(error.Message, EnumMessageLevel.Error, 0);

                    }
                    throw new Exception("Failed to read InputInfo csv !");
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.StackTrace);
            }
            finally
            {
                stopWatch.Stop();
                Response.Report(string.Format("Total Process Time : " + TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds).ToString(@"hh\:mm\:ss"), EnumMessageLevel.General));
            }
            return this;
        }
    }
}
