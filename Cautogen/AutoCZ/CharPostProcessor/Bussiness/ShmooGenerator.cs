using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.Controller;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.ShmooData;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.common.IgxlDataExtension;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using CharSetup = IgxlLib.IgxlBase.CharSetup;
using CharSheet = IgxlLib.IgxlSheets.CharSheet;
using CharStep = IgxlLib.IgxlBase.CharStep;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class ShmooGenerator
    {
        private const string SheetName = "DevChar_Sheet";
        private readonly bool _genCSharp = false;

        public ShmooGenerator(bool genCSharp = false)
        {
            _genCSharp = genCSharp;
        }

        public void Generate()
        {
            var seenGlobalSpec = new HashSet<string>();

            string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);

            var devCharSheet = new CharSheet(SheetName);
            LocalSpecs.TestProgram.JoblistSheet.AddCharSheet(devCharSheet.Name);
            foreach (ShmooSetup shmooSetup in LocalSpecs.AllShmooSetups)
            {
                shmooSetup.ShmooPins = IsPinGroupAndContainDCVI_DCVS(shmooSetup, out bool isPinGroup);
                BuildCharSetupForShmoo(shmooSetup, devCharSheet, seenGlobalSpec, isPinGroup);
            }
            string czFileName = Path.Combine(outputFolder, devCharSheet.Name + ".txt");
            devCharSheet.Write(czFileName, LocalSpecs.ExportVersion < 9.0 ? "2.5" : "2.6");
            LocalSpecs.GenSheets.Add(devCharSheet);
        }

        private static List<ShmooPin> IsPinGroupAndContainDCVI_DCVS(ShmooSetup shmooSetup, out bool isPinGroup)
        {
            List<ShmooPin> newShmooPins = [];
            isPinGroup = false;

            foreach (ShmooPin shmooPin in shmooSetup.ShmooPins)
            {
                if (LocalSpecs.ProgInfo.PinGroupDic.ContainsKey(shmooPin.SweepPinName.ToUpper()))
                {
                    List<string> groupPinList = LocalSpecs.ProgInfo.PinGroupDic[shmooPin.SweepPinName.ToUpper()].PinList
                        .Select(x => x.PinName.Replace("_", "").ToUpper()).ToList();
                    if (groupPinList.Any(x => x.Contains("DCVI", StringComparison.OrdinalIgnoreCase)) &&
                        groupPinList.Any(x => x.Contains("DCVS", StringComparison.OrdinalIgnoreCase)))
                    {
                        isPinGroup = true;

                        foreach (string pin in groupPinList)
                        {
                            if (pin.Contains("DCVI", StringComparison.OrdinalIgnoreCase))
                            {
                                newShmooPins.Add(new ShmooPin(
                                sweepPin: $"Vps:{pin}",
                                    start: shmooPin.StartPoint,
                                    stop: shmooPin.StopPoint,
                                    size: shmooPin.StepSize,
                                    shmooType: shmooPin.ShmooType)
                                {
                                    PortName = shmooPin.PortName,
                                });
                            }
                            else
                            {
                                newShmooPins.Add(new ShmooPin(
                                sweepPin: $"{shmooPin.SweepType}:{pin}",
                                    start: shmooPin.StartPoint,
                                    stop: shmooPin.StopPoint,
                                    size: shmooPin.StepSize,
                                    shmooType: shmooPin.ShmooType)
                                {
                                    PortName = shmooPin.PortName,
                                });
                            }
                        }
                    }
                    else
                    {
                        newShmooPins.Add(shmooPin);
                    }
                }
                else
                {
                    newShmooPins.Add(shmooPin);
                }
            }

            return newShmooPins;
        }

        private void BuildCharSetupForShmoo(ShmooSetup shmooSetup, CharSheet devCharSheet, HashSet<string> seenGlobalSpec, bool isPinGroup)
        {
            var charSetup = new CharSetup { SetupName = shmooSetup.ShmooSetupName, TestMethod = "Retest" };
            bool isXprimary = true;
            bool isYprimary = true;
            bool isZprimary = true;
            string allShmooTypes = string.Join("", shmooSetup.ShmooPins.Select(p => p.ShmooType));
            bool is3DShmoo = Regex.IsMatch(allShmooTypes, "X+Y+Z+", RegexOptions.IgnoreCase);

            foreach (ShmooPin shmooPin in shmooSetup.ShmooPins)
            {
                CharStep charStep = BuildCharStep(shmooSetup, shmooPin, seenGlobalSpec, isPinGroup);
                ApplyPrimaryStepConfig(charStep, shmooPin, shmooSetup, ref isXprimary, ref isYprimary, ref isZprimary);
                charStep.RangeFrom = shmooPin.StartPoint;
                charStep.RangeTo = shmooPin.StopPoint;
                if (is3DShmoo)
                {
                    Process3DShmooSetup(charStep);
                }

                charSetup.AddStep(charStep);
            }
            devCharSheet.AddRow(charSetup);
        }

        private CharStep BuildCharStep(ShmooSetup shmooSetup, ShmooPin shmooPin, HashSet<string> seenGlobalSpec, bool isPinGroup)
        {
            GetInterPosePointArg(shmooPin, shmooSetup.Timeset, out string preArgX, out string postArgX, out string preArgY, out string postArgY, out string preArgZ, out string postArgZ);

            var charStep = new CharStep(shmooSetup.ShmooSetupName, shmooPin.SweepPinName.Replace(",", "_"))
            {
                Mode = shmooPin.ShmooType + " Shmoo",
                PostStepArguments = ConstData.InterPostPostStepArg,
                PostStep = _genCSharp ? ConstData.InterPostPostStepCSharp : ConstData.InterPostPostStep,
                ApplyToPinExecMode = "Simultaneous",
                ApplyToPins = RemoveNcPins(shmooPin.SweepPinName, isPinGroup),
                OutputSuspendDatalog = string.IsNullOrEmpty(shmooSetup.SuspendDatalog)
                    ? "TRUE"
                    : shmooSetup.SuspendDatalog.ToUpper()
            };  //VMain:VDD_CPU,VDD_GPU

            if (shmooSetup.IsUseCmd)
            {
                charStep.PreSetup = "RTOS_Boot_CZ";
                charStep.PostPoint = "RTOS_Shmoo_Reboot";
                charStep.PostPointArguments = shmooSetup.ShmooSetupName;
            }

            if (Regex.IsMatch(shmooPin.SweepPinName, "^VDD", RegexOptions.IgnoreCase))
            {
                charStep.PostStepArguments += "," + charStep.ApplyToPins;
            }

            ApplyParameterName(charStep, shmooPin, seenGlobalSpec, preArgX, postArgX, preArgY, postArgY, preArgZ, postArgZ);
            charStep.ParameterType = GetParameterType(shmooPin);
            return charStep;
        }

        private void ApplyParameterName(CharStep charStep, ShmooPin shmooPin, HashSet<string> seenGlobalSpec,
            string preArgX, string postArgX, string preArgY, string postArgY, string preArgZ, string postArgZ)
        {
            if (Regex.IsMatch(shmooPin.SweepType, @"AC\s*Spec", RegexOptions.IgnoreCase) && shmooPin.PortName == "")
            {
                charStep.ParameterName = shmooPin.SweepPinName;
                return;
            }
            if (shmooPin.PortName == "")
            {
                if (Regex.IsMatch(shmooPin.SweepType, @"AC\s*Spec", RegexOptions.IgnoreCase))
                {
                    charStep.ParameterName = shmooPin.SweepPinName;
                    charStep.PrePoint = ConstData.InterPostPrePoint;
                    charStep.PostPoint = ConstData.InterPostPostPoint;
                    charStep.ParameterName = shmooPin.SweepPinName;
                    AssignAxisArguments(charStep, shmooPin.ShmooType, preArgX, postArgX, preArgY, postArgY, preArgZ, postArgZ);
                }
                else if (Regex.IsMatch(shmooPin.SweepType, @"Global\s*Spec", RegexOptions.IgnoreCase))
                {
                    string pins = charStep.ApplyToPins.Replace(",", "_").Replace(" ", "");
                    charStep.ParameterName = "Shmoo_" + "Glb" + "_GLB";
                    AddToGlobalSpec(seenGlobalSpec, charStep.ParameterName);
                }
                else
                {
                    charStep.ParameterName = shmooPin.SweepType;
                }
                return;
            }

            charStep.PrePoint = ConstData.InterPostPrePoint;
            charStep.PostPoint = ConstData.InterPostPostPoint;
            charStep.ParameterName = shmooPin.SweepPinName;
            AssignAxisArguments(charStep, shmooPin.ShmooType, preArgX, postArgX, preArgY, postArgY, preArgZ, postArgZ);
        }

        private static void AssignAxisArguments(CharStep charStep, string shmooType, string preArgX, string postArgX, string preArgY, string postArgY, string preArgZ, string postArgZ)
        {
            switch (shmooType)
            {
                case "X":
                    charStep.PrePointArguments = preArgX;
                    charStep.PostPointArguments = postArgX;
                    break;

                case "Y":
                    charStep.PrePointArguments = preArgY;
                    charStep.PostPointArguments = postArgY;
                    break;

                case "Z":
                    charStep.PrePointArguments = preArgZ;
                    charStep.PostPointArguments = postArgZ;
                    break;
            }
        }

        private static void ApplyPrimaryStepConfig(CharStep charStep, ShmooPin shmooPin, ShmooSetup shmooSetup, ref bool isXprimary, ref bool isYprimary, ref bool isZprimary)
        {
            if ((shmooPin.ShmooType == "X" && isXprimary) || (shmooPin.ShmooType == "Y" && isYprimary)
                || (shmooPin.ShmooType == "Z" && isZprimary))
            {
                charStep.RangeCalcField = "steps";
                charStep.RangeStepSize = shmooPin.StepSize;
                charStep.AlgorithmName = shmooSetup.SearchMethod;
                charStep.AlgorithmName = shmooPin.ShmooType == "Y" || shmooPin.ShmooType == "Z" ? "Linear" : shmooSetup.SearchMethod;
                charStep.OutputFormat = "Enhanced";
                charStep.OutputDestinationsTextFile = "Disable";
                charStep.OutputDestinationsSheet = "Disable";
                charStep.OutputDestinationsDatalog = "Enable";
                charStep.OutputDestinationsImmediateWin = "Disable";
                charStep.OutputDestinationsOutputWin = "Disable";

                if (shmooSetup.SearchMethod.ToLower() == "jump")
                {
                    charStep.AlgorithmArguments = "6";
                }
            }
            switch (shmooPin.ShmooType)
            {
                case "X":
                    isXprimary = false;
                    break;

                case "Y":
                    isYprimary = false;
                    break;

                case "Z":
                    isZprimary = false;
                    break;
            }
        }

        private string GetParameterType(ShmooPin shmooPin)
        {
            string parameterType = "Level";

            if (Regex.IsMatch(shmooPin.SweepType, @"AC\s*Spec", RegexOptions.IgnoreCase) || shmooPin.PortName != "")
            {
                parameterType = "AC Spec";
            }
            else if (Regex.IsMatch(shmooPin.SweepType, @"DC\s*Spec", RegexOptions.IgnoreCase))
            {
                parameterType = "DC Spec";
            }
            else if (Regex.IsMatch(shmooPin.SweepType, @"Global\s*Spec", RegexOptions.IgnoreCase))
            {
                parameterType = "Global Spec";
            }

            return parameterType;
        }

        private void AddToGlobalSpec(ISet<string> seenGlobalSpec, string charStepParameterName)
        {
            if (seenGlobalSpec.Contains(charStepParameterName))
            {
                return;
            }

            seenGlobalSpec.Add(charStepParameterName);

            if (LocalSpecs.TestProgram.GlobalSpecRows.Exists(
                p => p.Symbol.Equals(charStepParameterName, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            LocalSpecs.TestProgram.GlobalSpecRows.Add(
                new GlobalSpec(charStepParameterName, "0"));
        }

        private void GetInterPosePointArg(ShmooPin shmooPin, string timesetSheetName,
            out string preArgX, out string postArgX, out string preArgY,
            out string postArgY, out string preArgZ, out string postArgZ)
        {
            // reset args to empty string
            preArgX = "";
            postArgX = "";
            preArgY = "";
            postArgY = "";
            preArgZ = "";
            postArgZ = "";

            // try to find the corresponding port setting in timeSet sheet, else use default value
            TimeSetBasicSheet timesetSheet = LocalSpecs.TestProgram.TimeSetSheets
                .FirstOrDefault(p => p.Name.Equals(timesetSheetName, StringComparison.OrdinalIgnoreCase));

            TSet timesetRow = null;

            var portNameList =
                LocalSpecs.TestProgram.PortRows.Where(x => !string.IsNullOrEmpty(x.PortName))
                    .Select(x => x.PortName.Replace("_Port", ""))
                    .Distinct()
                    .ToList();

            if (!Regex.IsMatch(shmooPin.SweepType, @"AC\s*Spec", RegexOptions.IgnoreCase))
            {
                return;
            }

            string portName = portNameList.FirstOrDefault(x => Regex.IsMatch(x.Replace("_", ""), shmooPin.SweepPinName.Replace("_", ""), RegexOptions.IgnoreCase));
            if (string.IsNullOrEmpty(portName) && Regex.IsMatch(shmooPin.SweepPinName, "^XI0_Diff", RegexOptions.IgnoreCase))
            {
                shmooPin.PortName = ConstData.InterPostDefaultPort;
            }

            if (timesetSheet != null)
            {
                foreach (TSet timeset in timesetSheet.Rows
                    .Where(timeset => timeset.TimingRows.Exists(a => a.PinGrpName.Split('_')[0].Equals(shmooPin.SweepPinName)
                                                                     && a.DataSrc.ToLower() == "pa")))
                {
                    timesetRow = timeset;
                    break;
                }
            }

            if (timesetRow != null)
            {
                string shmooPar = Regex.IsMatch(timesetRow.CyclePeriod, @"_\w+")
                    ? Regex.Match(timesetRow.CyclePeriod, @"_(?<par>\w+)").Groups["par"].Value
                    : $"CyclePeriodHardCode:{timesetRow.CyclePeriod}";
                shmooPin.PortName = timesetRow.Name;
                shmooPin.SweepPinName = shmooPar;
            }
            //else
            //{
            //    shmooPin.PortName = portName + "_Port";
            //    shmooPin.SweepPinName = portName + "_Freq_VAR";
            //}

            switch (shmooPin.ShmooType)
            {
                case "X":
                    preArgX = shmooPin.ShmooType + "," + shmooPin.PortName + "," + shmooPin.SweepPinName;
                    postArgX = shmooPin.PortName;
                    break;

                case "Y":
                    preArgY = shmooPin.ShmooType + "," + shmooPin.PortName + "," + shmooPin.SweepPinName;
                    postArgY = shmooPin.PortName;
                    break;

                case "Z":
                    preArgZ = shmooPin.ShmooType + "," + shmooPin.PortName + "," + shmooPin.SweepPinName;
                    postArgZ = shmooPin.PortName;
                    break;
            }
        }

        private string RemoveNcPins(string applyPins, bool isPingGroup)
        {
            List<string> pinList = applyPins.Replace("_", "").Split(',').ToList(); // Some CharPlan may use "_" in pin name

            if (ProdProg.ChannelMapSheet == null)
            {
                return applyPins;
            }

            var newPinList = new List<string>();
            foreach (string applyPin in pinList)
            {
                ChannelMapRow pin = ProdProg.ChannelMapSheet.Rows.FirstOrDefault(a => a.DeviceUnderTestPinName.Replace("_", "")
                    .Equals(applyPin, StringComparison.OrdinalIgnoreCase));
                if (pin != null)
                {
                    if (isPingGroup || (pin.Type.ToLower() != "n/c" && pin.Type.ToLower() != "utility"))
                    {
                        newPinList.Add(pin.DeviceUnderTestPinName);
                    }
                }
                else
                {
                    if (LocalSpecs.ProgInfo.PinDic.ContainsKey(applyPin.ToUpper()))
                    {
                        newPinList.Add(LocalSpecs.ProgInfo.PinDic[applyPin.ToUpper()]);
                    }
                }
            }
            return string.Join(",", newPinList);
        }

        private void Process3DShmooSetup(CharStep charStep)
        {
            charStep.AxisExecutionOrder = "Z-X-Y";
            charStep.PreSetup = "StoreMaxNum";
        }
    }
}
