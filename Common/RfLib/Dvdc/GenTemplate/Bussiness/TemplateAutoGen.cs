using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using RfLib.Dvdc.GenTemplate.TestPlanFormat;
using RfLib.Dvdc.Reader.CapturePostProcess;
using RfLib.InstrumentSetup;

using TestPlanLib.Basic.Relay;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public partial class TemplateAutoGen
    {
        [GeneratedRegex("WiMeas|WiSrc", RegexOptions.IgnoreCase, "en-US")]
        internal static partial Regex MyRegex3();
        [GeneratedRegex("rffunc_trx_universal", RegexOptions.IgnoreCase, "en-US")]
        internal static partial Regex MyRegex4();
        [GeneratedRegex("sweep", RegexOptions.IgnoreCase, "en-US")]
        internal static partial Regex MyRegex11();

        public List<InstrumentSetupRow> InstrumentSetupForPatList = [];
        public List<SetEfuseItem> ReferenceSetEfuseItems = [];
        public Dictionary<string, List<PostProcessSheetRow>> Cpps = [];
        private readonly ChannelMapPinResolver _pinResolver;

        internal HardIpInputData HardIpInputDataInternal { get; }

        internal List<ChannelMapSheet> MultiChannelMapInternal => _pinResolver.MultiChannelMap;

        internal RelayTableNew RelayInfoInternal { get; }
        internal List<string> DsscSetupNameTestPlanInternal { get; } = [];
        internal PatternRowModifier PatternRowModifierInternal { get; }
        internal WiTrimBestCodeGenerator WitrimBestCodeGeneratorInternal { get; }
        internal PlanGenerator PlanGeneratorInternal { get; }
        internal RegisterLimitationGenerator RegisterLimitationGeneratorInternal { get; }

        public TemplateAutoGen(HardIpInputData hardIpInputData, List<ChannelMapSheet> channelMapSheets, RelayTableNew relayTableNew)
        {
            HardIpInputDataInternal = hardIpInputData;
            _pinResolver = new ChannelMapPinResolver(channelMapSheets);
            RelayInfoInternal = relayTableNew;
            PatternRowModifierInternal = new PatternRowModifier(this);
            WitrimBestCodeGeneratorInternal = new WiTrimBestCodeGenerator(this);
            PlanGeneratorInternal = new PlanGenerator(this);
            RegisterLimitationGeneratorInternal = new RegisterLimitationGenerator(this);
            string basicConfig = string.Format("{0}\\Settings\\Basic\\Basic_Configure_{1}.xlsx", LocalSpecs.SettingFolder, LocalSpecs.CurrentProject);
            if (File.Exists(basicConfig))
            {
                var inputExcel = new ExcelPackage(new FileInfo(basicConfig));
                SettingStatic.BasicConfigWorkbook = inputExcel.Workbook;
                string outString = string.Format("Load Basic_Configure_{0}.xlsx Successfully!!", LocalSpecs.CurrentProject);
                Response.Report(outString, percentage: Convert.ToInt32(100));
            }
            else
            {
                string outString = string.Format("Basic_Configure_{0}.xlsx Not Found in folder \"Settings\"??", LocalSpecs.CurrentProject);
                Response.Report(outString, EnumMessageLevel.Error, Convert.ToInt32(100));
            }
        }

        internal void GenPlanWithNonVdiffNew(string payload, List<TemplateRow> templateRows, HardIpSeqInfoNew hardIpSeqInfoNew, string seqindex, int blockindex, int stepIndex, bool isRFItem, bool isBBItem = false)
        {
            PlanGeneratorInternal.GenPlanWithNonVdiffNew(payload, templateRows, hardIpSeqInfoNew, seqindex, blockindex, stepIndex, isRFItem, isBBItem);
        }

        internal string SearchPinInChannelMap(MeasPin measPin)
        {
            return _pinResolver.SearchPinInChannelMap(measPin);
        }
    }
}
