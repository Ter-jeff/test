using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.Settings;
using TestPlanLib.Static;
using TestPlanLib.Xml;

namespace Automation.GenerateIgxl.PreAction.GenPinMap
{
    public class PinMapMain
    {
        private readonly bool _needNwire = true;

        public PinMapSheet WorkFlow(PinMapSheet pinMapSheet, ExcelWorksheet ioPinGroupSheet, IoContiSheet ioContiSheet)
        {
            if (pinMapSheet == null)
            {
                throw new Exception("Missing " + NeededSheets.PinMap + " sheet in test plan file");
            }

            PinMapSheet sheet = WorkFlow_PinMapSheet(pinMapSheet, ioPinGroupSheet, ioContiSheet);
            return sheet;
        }

        private PinMapSheet WorkFlow_PinMapSheet(PinMapSheet pinMapSheet, ExcelWorksheet ioPinGroupSheet, IoContiSheet ioContiSheet)
        {
            PinMapSheet wSheet = pinMapSheet ?? throw new Exception(" Missing " + NeededSheets.PinMap + " sheet in test plan file");

            if (LocalSpecs.Options.Device != EnumDevice.LCD && ioContiSheet == null)
            {
                throw new Exception($" Missing {NeededSheets.ContiIo} sheet in test plan file");
            }

            return CreatePinMap(wSheet, ioPinGroupSheet, ioContiSheet);
        }

        private PinMapSheet CreatePinMap(PinMapSheet pinMapSheet, ExcelWorksheet ioPinGroupSheet, IoContiSheet ioContiSheet)
        {
            if (pinMapSheet == null)
            {
                return null;
            }

            ExcelWorksheet ioLevelSheet = EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.IoLevels];
            var ioIgnoreListSheet = new IoIgnoreList(EpWorkbook.TestPlanWorkbook.Worksheets.FirstOrDefault(p => Regex.IsMatch(p.Name, "IO_ignore_list", RegexOptions.IgnoreCase)));
            var ioConti = IoContiSheet.GenIoContiSheet(ioContiSheet, ioLevelSheet);

            IoContiSheet.CheckIoPins(ioConti, ioIgnoreListSheet.IoIgnorePinsTypes, pinMapSheet);

            if (ioPinGroupSheet != null)
            {
                List<PinGroup> groupFromPinGropSheet = ReadGroupSheet(pinMapSheet, ioPinGroupSheet);
                MergePinMapSheet(pinMapSheet, groupFromPinGropSheet, $"{NeededSheets.IoGroup} sheet");
            }

            //Add group like "Pins_1p1v"
            List<PinGroup> ioLevelGroup = GenerateIoLevelGroup(pinMapSheet, ioConti);
            MergePinMapSheet(pinMapSheet, ioLevelGroup, $"{NeededSheets.ContiIo} sheet");

            List<PinGroup> digitalGroup = GenDigitalGroup(pinMapSheet, ioConti);
            MergePinMapSheet(pinMapSheet, digitalGroup);               //Add All_digital group

            if (!pinMapSheet.IsGroupExist("CONTINUITY_NEG_AUTOZ"))
            {
                PinGroup contiNegAutoZPinGroup = GenContiNegGroup(pinMapSheet, ioConti);
                MergePinMapSheet(pinMapSheet, new List<PinGroup> { contiNegAutoZPinGroup }, "all I/O pins in IO_Continuity sheet");      //Add CONTINUITY_NEG_AUTOZ group
            }

            if (LocalSpecs.Options.Device == EnumDevice.AP)
            {
                List<PinGroup> spiGroups = GenSpiGroups(pinMapSheet);                 //Add SPIROM_PINS group
                MergePinMapSheet(pinMapSheet, spiGroups);
            }

            //Add nWire group
            if (_needNwire)
            {
                List<PinGroup> nWireGroup = GenerateNWireGroups(ioConti);
                MergePinMapSheet(pinMapSheet, nWireGroup, "FreeRunningClock sheet");
            }

            return pinMapSheet;
        }

        private List<PinGroup> GenDigitalGroup(PinMapSheet pinMapSheet, IoContiSheet ioConti)
        {
            var pinGroups = new List<PinGroup>();
            List<Pin> pinList = pinMapSheet.GetIoPins();

            var digitalList = pinList.Select(a => a.PinName).ToList();
            var pinGroup = new PinGroup("All_Digital", PinMapConst.TypeIo);
            pinGroup.AddPins(digitalList, PinMapConst.TypeIo);
            pinGroups.Add(pinGroup);

            var digitalDiscList = pinList.Where(a => !Regex.IsMatch(a.PinName, ".*_PA$|^REFCLK_.*", RegexOptions.IgnoreCase))
                .Select(a => a.PinName).ToList();
            pinGroup = new PinGroup("All_Digital_Disc", PinMapConst.TypeIo);
            pinGroup.AddPins(digitalDiscList, PinMapConst.TypeIo);
            pinGroups.Add(pinGroup);

            List<string> pinListFromIoConti = ioConti.GetPinList();
            var digitalPowerUpList = pinList.Where(a => pinListFromIoConti.Exists(p => p.Equals(a.PinName, StringComparison.OrdinalIgnoreCase)))
                .Select(a => a.PinName).ToList();
            pinGroup = new PinGroup("All_Digital_PowerUp", PinMapConst.TypeIo);
            pinGroup.AddPins(digitalPowerUpList, PinMapConst.TypeIo);
            pinGroups.Add(pinGroup);

            pinGroup = new PinGroup("All_DiffPairs", "I/O");
            Dictionary<string, string> pairs = DifferentialPair(pinList.Select(p => p.PinName).ToList());
            foreach (KeyValuePair<string, string> keyValuePair in pairs)
            {
                if (!pinGroup.PinList.Select(x => x.PinName).Contains(keyValuePair.Key))
                {
                    pinGroup.AddPin(keyValuePair.Key, PinMapConst.TypeIo);
                    pinGroup.AddPin(keyValuePair.Value, PinMapConst.TypeIo);
                }
            }
            pinGroups.Add(pinGroup);
            return pinGroups;
        }

        private PinGroup GenContiNegGroup(PinMapSheet pinMapSheet, IoContiSheet ioConti)
        {
            List<Pin> pinList = pinMapSheet.GetIoPins();

            List<string> pinListFromIoConti = ioConti.GetPinList();
            var allIoList = pinList.Where(a => pinListFromIoConti.Exists(p => p.Equals(a.PinName, StringComparison.OrdinalIgnoreCase) && a.PinType == "I/O"))
                .Select(a => a.PinName).ToList();
            var pinGroup = new PinGroup("CONTINUITY_NEG_AUTOZ", PinMapConst.TypeIo);
            pinGroup.AddPins(allIoList, PinMapConst.TypeIo);


            return pinGroup;
        }

        private static Dictionary<string, string> DifferentialPair(List<string> pinList)
        {
            DiffPairConfig config = XmlService<DiffPairConfig>.LoadXml(Path.Combine(AppContext.BaseDirectory, "Config", "DiffPairConfig.xml"));
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            foreach (DiffItem pinPair in config.DiffPairPins)
            {
                //Add differential pairs from config which defined using Pin name
                string posPin = pinList.Find(p => p.Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase));
                string negPin = pinList.Find(p => p.Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase));
                if (posPin != null && negPin != null && !pairs.ContainsKey(posPin))
                {
                    pairs.Add(posPin, negPin);
                }
            }

            for (int i = 0; i < pinList.Count; i++)
            {
                for (int j = i + 1; j < pinList.Count; j++)
                {
                    if (pinList[i].Length == pinList[j].Length)
                    {
                        GetSamePartInDiffPairs(pinList[i], pinList[j], out string nStr, out string pStr);
                        bool flag = false;
                        foreach (DiffItem rule in config.DiffPairRules)
                        {
                            if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.ContainsKey(pinList[j]))
                                {
                                    pairs.Add(pinList[j], pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                            else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.ContainsKey(pinList[i]))
                                {
                                    pairs.Add(pinList[i], pinList[j]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                        }
                        if (flag)
                        {
                            break;
                        }
                    }
                }
            }

            //Store Neg and Pos pin again
            int count = pairs.Count;
            for (int i = 0; i < count; i++)
            {
                pairs.Add(pairs.ElementAt(i).Value, pairs.ElementAt(i).Key);
            }
            return pairs;
        }

        private static void GetSamePartInDiffPairs(string nPinName, string pPinName, out string nStr, out string pStr)
        {
            //EX:ADDR_M2P_DQ_N0::ADDR_M2P_DQ_P0
            pStr = "";
            nStr = "";
            if (nPinName.Length == pPinName.Length)
            {
                for (int i = 0; i < nPinName.Length; i++)
                {
                    if (nPinName[i] != pPinName[i])
                    {
                        nStr += nPinName[i];
                        pStr += pPinName[i];
                    }
                }
            }
        }

        private List<PinGroup> ReadGroupSheet(PinMapSheet pinMap, ExcelWorksheet pinGroupSheet)
        {
            var pinGroups = new List<PinGroup>();
            if (pinGroupSheet.Dimension == null)
            {
                return pinGroups;
            }
            int format = CheckIoPinGroupSheetFormat(pinGroupSheet);
            if (format == 0)
            {
                pinGroups = ParsePinGroupSheet(pinMap, pinGroupSheet);
            }
            else if (format == 1)
            {
                pinGroups = new IoPinGroupReaderNew().ReadMain(pinGroupSheet);
            }

            pinGroups = ModifyPinGroups(pinMap, pinGroups);
            return pinGroups;
        }

        private List<PinGroup> ParsePinGroupSheet(PinMapSheet pinMap, ExcelWorksheet pinGroupSheet)
        {
            var pinGroups = new List<PinGroup>();
            int endRow = pinGroupSheet.Dimension.End.Row;
            int startRow = -1;
            int startCol = -1;
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (Regex.IsMatch(EpplusExtensions.GetCellValue(pinGroupSheet, i, j), "Group", RegexOptions.IgnoreCase))
                    {
                        startRow = i;
                        startCol = j;
                        break;
                    }
                }
                if (startCol != -1)
                {
                    break;
                }
            }
            if (startCol == -1)
            {
                return pinGroups;
            }

            for (int i = startRow + 1; i <= endRow;)
            {
                string groupName = EpplusExtensions.GetCellValue(pinGroupSheet, i, startCol).Trim();
                if (groupName.Equals(""))
                {
                    i++;
                    continue;
                }

                var pins = new List<Pin>();
                while (i <= endRow && (EpplusExtensions.GetCellValue(pinGroupSheet, i, startCol).Trim().Equals("") || EpplusExtensions.GetCellValue(pinGroupSheet, i, startCol).Trim().Equals(groupName, StringComparison.OrdinalIgnoreCase)))
                {
                    string pinName = EpplusExtensions.GetCellValue(pinGroupSheet, i, startCol + 1).Trim();
                    if (pinName.Contains("+") || pinName.Contains("-"))
                    {
                        pins.AddRange(PinGroupOperation(pinMap, pinName));
                    }
                    else if (pinName.Contains("*"))
                    {
                        pins.AddRange(SearchPins(pinMap, pinName));
                    }
                    else
                    {
                        pinName = pinName.Replace("[", "").Replace("]", "");
                        if (!pinName.Equals(""))
                        {
                            if (pinMap.IsGroupExist(pinName))
                            {
                                pins.Add(new Pin(pinName, pinMap.GetGroup(pinName).PinType));
                            }
                            else if (pinGroups.Exists(p => p.PinName.ToLower().Equals(pinName.ToLower())))
                            {
                                PinGroup target = pinGroups.FirstOrDefault(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                                pins.Add(target != null ? new Pin(pinName, target.PinType) : new Pin(pinName, pinGroups.Select(p => p.PinType).FirstOrDefault()));
                            }
                            else
                            {
                                if (pinMap.IsPinExist(pinName))
                                {
                                    pins.Add(pinMap.GetPin(pinName));
                                }
                                else
                                {
                                    var newPin = new Pin(pinName, "");
                                    pins.Add(newPin);
                                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_01, pinGroupSheet.Name, i, startCol + 1, $"The pin {pinName} in IO_PinGroup can not be found !!!", new string[] { pinName });
                                }
                            }
                        }
                    }
                    i++;
                    //groupName = EpplusExtensions.GetCellValue(pinGroupSheet, i, startCol).Trim();
                }

                var pinTypes = pins.Select(x => x.PinType).Distinct().ToList();
                string pinType = pinTypes.Count == 1 ? pinTypes.First() : "";
                if (pinTypes.Count > 1)
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_02, pinGroupSheet.Name, i, startCol, $"The pin group {groupName} has more than two pin types !!!", new string[] { groupName });
                }
                var group = new PinGroup(groupName, pinType);
                if (pins.Any(x => x.PinName.Equals(groupName)))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_03, pinGroupSheet.Name, i, startCol, $"The pin group name {groupName} and pin name can not be the same !!!", new string[] { groupName });
                }
                pins.ForEach(p => p.Comment = pinGroupSheet.Name);
                group.AddPins(pins);
                pinGroups.Add(group);

            }
            return pinGroups;
        }

        private List<Pin> PinGroupOperation(PinMapSheet pinMap, string operationCommand)
        {
            string lStrPattern = "[+|-]";
            List<string> keyWordList = Regex.Split(operationCommand, lStrPattern).ToList();
            var opList = Regex.Matches(operationCommand, lStrPattern).Cast<Match>().Select(a => a.Value).ToList();

            List<Pin> pins = SearchPins(pinMap, keyWordList.First());
            for (int i = 0; i < keyWordList.Count - 1; i++)
            {
                List<Pin> pinListTemp = SearchPins(pinMap, keyWordList[i + 1]);
                if (opList[i].Equals("+"))
                {
                    //Union
                    foreach (Pin pin in pinListTemp)
                    {
                        if (!pins.Exists(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            pins.Add(pin);
                        }
                    }
                }
                else if (opList[i].Equals("-"))
                {
                    //Intersection
                    foreach (Pin pin in pinListTemp)
                    {
                        if (pins.Exists(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            pins.RemoveAll(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase));
                        }
                    }
                }
                else
                {
                    throw new Exception($"Unknow operation {opList[i]}");
                }
            }

            return pins;
        }

        private List<Pin> SearchPins(PinMapSheet pinMap, string keyWord)
        {
            var pins = new List<Pin>();
            if (keyWord.Contains("*"))
            {
                //use greed match method
                string lStrPattern = keyWord.Replace("*", ".*");
                pins.AddRange(pinMap.PinList.Where(p => Regex.IsMatch(p.PinName, lStrPattern, RegexOptions.IgnoreCase)));
            }
            else if (pinMap.IsGroupExist(keyWord))
            {
                PinGroup group = pinMap.GetGroup(keyWord);
                foreach (Pin pin in group.PinList)
                {
                    pins.AddRange(SearchPins(pinMap, pin.PinName));
                }

            }
            else if (pinMap.IsPinExist(keyWord))
            {
                //if it is a pin
                pins.Add(pinMap.GetPin(keyWord));
            }
            else
            {
                throw new Exception($"Unknown command {keyWord}");
            }
            return pins;
        }

        private void MergePinMapSheet(PinMapSheet pinMapSheet, List<PinGroup> pinGroupList, string sourceSheet = "AutogenRule")
        {
            foreach (PinGroup pinGroup in pinGroupList)
            {
                if (!pinGroup.PinList.Any())
                {
                    continue;
                }

                if (pinMapSheet.IsGroupExist(pinGroup.PinName))
                {
                    PinGroup mapGroup = pinMapSheet.GroupList.FirstOrDefault(a => a.PinName.Equals(pinGroup.PinName, StringComparison.OrdinalIgnoreCase));
                    if (mapGroup != null && (!mapGroup.PinList.All(a => pinGroup.PinList.Select(p => p.PinName.ToUpper()).ToList().Contains(a.PinName.ToUpper())) || !pinGroup.PinList.All(a => mapGroup.PinList.Select(p => p.PinName.ToUpper()).ToList().Contains(a.PinName.ToUpper()))))
                    {
                        ErrorReportManager.AddError(PreActionErrorType.E_MismatchPinGroup_01, "", 1, 0,
                            [pinGroup.PinName, sourceSheet, string.Join(",", pinGroup.PinList.Select(x => x.PinName).ToArray())]);

                        //merge the pin group if there is some pins not grouped in the pin map
                        if (!pinGroup.PinList.All(mapGroup.PinList.Contains))
                        {
                            var pins = pinGroup.PinList.Select(p => p).Where(p => !mapGroup.PinList
                                .Select(x => x.PinName.ToUpper()).Contains(p.PinName.ToUpper())).ToList();
                            mapGroup.PinList.AddRange(pins);
                        }

                    }
                    if (mapGroup != null && pinGroup.PinList.First().PinType != mapGroup.PinList.First().PinType && !string.IsNullOrEmpty(pinGroup.PinList.First().PinType))
                    {
                        ErrorReportManager.AddError(PreActionErrorType.E_MismatchPinGroup_02, "", 1, 0,
                            [mapGroup.PinName, sourceSheet, pinGroup.PinList.First().PinType, mapGroup.PinList.First().PinType]);
                    }
                    continue;
                }
                pinMapSheet.AddRow(pinGroup);
            }
        }

        private List<PinGroup> GenerateIoLevelGroup(PinMapSheet pinMap, IoContiSheet ioContiSheet)
        {
            var pinGroups = new List<PinGroup>();
            var lLvlCategory = new Dictionary<string, List<string>>();
            List<string> pinListFromIoConti = ioContiSheet.GetPinList();
            List<EnumEquipment> equipments = TestPlanStatic.Equipments;
            var nWirePins = new List<string>();
            foreach (EnumEquipment equipment in equipments)
            {
                nWirePins.AddRange(NwireSingleton.Instance().SettingInfo.NwirePins.Where(x => x.PinType != IoPinType.Diff).Select(x => x.CreatePaClkPinName(equipment)).ToList());
            }
            foreach (string pinInIoConti in pinListFromIoConti)
            {
                string voltage = ioContiSheet.GetVoltage(pinInIoConti);
                voltage = voltage.Replace("V", "").Replace("v", "").Replace("m", "").Replace(" ", "").Replace("\t", "");
                if (voltage.Length > 0)
                {
                    foreach (Pin pin in pinMap.PinList)
                    {
                        if (pin.PinType.Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase)
                            || pin.PinType.Equals("I", StringComparison.OrdinalIgnoreCase) ||
                            pin.PinType.Equals("O", StringComparison.OrdinalIgnoreCase))
                        {
                            if (pin.PinName.Equals(pinInIoConti, StringComparison.OrdinalIgnoreCase))
                            {
                                if (!lLvlCategory.ContainsKey(voltage))
                                {
                                    lLvlCategory[voltage] = new List<string>();
                                }

                                if (!nWirePins.Exists(x => x.Equals(pinInIoConti, StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    lLvlCategory[voltage].Add(pinInIoConti);
                                }
                            }
                        }
                    }
                }
            }
            foreach (string key in lLvlCategory.Keys)
            {
                string pinGrpName;
                string pinVol = key.Split('_').First();
                string chiplet = Regex.Match(key, @"\w+_(?<chiplet>[A-z]\d+$)").Groups["chiplet"].ToString();
                if (pinVol.Equals("0"))
                {
                    continue;
                }

                if (pinVol.Contains("."))
                {
                    pinGrpName = Combination.CombineByUnderLine("Pins_" + pinVol.Replace(".", "p") + "v", chiplet);
                    string name = pinGrpName;
                    pinGrpName = Regex.Replace(pinGrpName, @"(?<voltage>\d+[.]\d+)[0]+", delegate
                    {
                        string num = Regex.Match(name, @"(?<voltage>\d+[.]\d+)[0]+").Groups["voltage"].ToString();
                        return num;
                    });
                }
                else
                {
                    pinGrpName = Combination.CombineByUnderLine("Pins_" + pinVol + "p0v", chiplet);
                }

                var lGroup = new PinGroup(pinGrpName, PinMapConst.TypeIo);
                foreach (string pin in lLvlCategory[key].Distinct())
                {
                    var newPin = new Pin(pin, PinMapConst.TypeIo, "IO sheet");
                    lGroup.AddPin(newPin);
                }
                pinGroups.Add(lGroup);
            }

            return pinGroups;
        }

        private List<PinGroup> GenerateNWireGroups(IoContiSheet ioConti)
        {
            var nWireGroups = new List<PinGroup>();
            List<ProtocolAwarePin> nWirePins = NwireSingleton.Instance().SettingInfo.NwirePins;
            var nWirePosGroup = new PinGroup(CommonConst.NWirePosGroup, "I/O");
            var nWireNegGroup = new PinGroup(CommonConst.NWireNegGroup, "I/O");
            List<EnumEquipment> equipments = TestPlanStatic.Equipments;
            foreach (EnumEquipment equipment in equipments)
            {
                foreach (ProtocolAwarePin nWirePin in nWirePins)
                {
                    string value;
                    if (ioConti.IsPinExist(nWirePin.OutClk))
                    {
                        ioConti.TryGetFsDd(nWirePin.OutClk, out value);
                    }
                    else
                    {
                        ioConti.TryGetFsDd(nWirePin.CreatePaClkPinName(equipment), out value);
                    }

                    if (value.Equals("DD", StringComparison.OrdinalIgnoreCase))
                    {
                        nWirePosGroup.AddPin(nWirePin.CreatePaClkPinName(equipment));
                    }
                    nWireNegGroup.AddPin(nWirePin.CreatePaClkPinName(equipment));

                    PinGroup group;
                    if (nWirePin.PinType == IoPinType.Diff)
                    {
                        //Add differential pin group, Out clock pin and out differential clock pin
                        group = new PinGroup(nWirePin.CreateDiffPinGroupName(equipment), "Differential");
                        group.AddPin(nWirePin.CreatePaClkPinName(equipment), "Differential");
                        group.AddPin(nWirePin.CreatePaClkDiffPinName(equipment), "Differential");
                        nWireGroups.Add(group);

                        //Add port map pin group, out clock pin and Reference Clock Pin
                        group = new PinGroup(nWirePin.CreatePortName(equipment), "I/O");
                        group.AddPin(nWirePin.CreatePaClkPinName(equipment));
                        if (equipment != EnumEquipment.UltraFlexPlus)
                        {
                            group.AddPin(nWirePin.RefClk);
                        }

                        group.AddPin(nWirePin.CreatePaClkDiffPinName(equipment));
                        if (!string.IsNullOrEmpty(nWirePin.ExtraPin))
                        {
                            const string regPin = @"^(?<pin>[\w]+)[\s]*([\(](?<name>[^)]+)[\)])?";
                            foreach (string item in nWirePin.ExtraPin.Split(','))
                            {
                                string pin =
                                    Regex.Match(item, regPin, RegexOptions.IgnoreCase).Groups["pin"].ToString().Trim();
                                group.AddPin(pin);
                            }

                        }
                        nWireGroups.Add(group);

                        //Add XO0 to Nwire pos or neg group
                        if (ioConti.IsPinExist(nWirePin.OutClkDiff))
                        {
                            ioConti.TryGetFsDd(nWirePin.OutClkDiff, out value);
                        }
                        else
                        {
                            ioConti.TryGetFsDd(nWirePin.CreatePaClkDiffPinName(equipment), out value);
                        }

                        if (value == "")
                        {

                        }
                        else if (value.Equals("DD", StringComparison.OrdinalIgnoreCase))
                        {
                            nWirePosGroup.AddPin(nWirePin.CreatePaClkDiffPinName(equipment));
                        }
                        nWireNegGroup.AddPin(nWirePin.CreatePaClkDiffPinName(equipment));
                    }
                    else
                    {
                        //Add port map pin group, out clock pin and Reference Clock Pin
                        group = new PinGroup(nWirePin.CreatePortName(equipment), "I/O");
                        group.AddPin(nWirePin.CreatePaClkPinName(equipment));
                        if (!string.IsNullOrEmpty(nWirePin.ExtraPin))
                        {
                            const string regPin = @"^(?<pin>[\w]+)[\s]*([\(](?<name>[^)]+)[\)])?";
                            foreach (string item in nWirePin.ExtraPin.Split(','))
                            {
                                string pin = Regex.Match(item, regPin, RegexOptions.IgnoreCase).Groups["pin"].ToString().Trim();
                                group.AddPin(pin);
                            }
                        }
                        if (equipment != EnumEquipment.UltraFlexPlus)
                        {
                            group.AddPin(nWirePin.RefClk);
                        }

                        if (group.PinList.Count > 1)
                        {
                            nWireGroups.Add(group);
                        }
                    }
                }

                nWireGroups.Add(nWireNegGroup);
                nWireGroups.Add(nWirePosGroup);
            }
            List<NonFrcNWires> nonFrcNWiresList = NwireSingleton.Instance().NonFrcSetting;
            if (nonFrcNWiresList != null)
            {
                var list = nonFrcNWiresList.GroupBy(x => x.PortName).ToList();
                foreach (IGrouping<string, NonFrcNWires> set in list)
                {
                    var nonFrc = new PinGroup(set.Key, "I/O");
                    foreach (NonFrcNWires item in set)
                    {
                        nonFrc.AddPin(item.DeviecPinName);
                    }
                    nWireGroups.Add(nonFrc);
                }
            }
            return nWireGroups;
        }

        private List<PinGroup> GenSpiGroups(PinMapSheet pinMapSheet)
        {
            var pinGroups = new List<PinGroup>();
            const string regexPattern = "^SPI0_.*";
            const string regexPattern1 = "^RTOS0_.*";
            if (pinMapSheet.IsGroupExist(CommonConst.SpiRomPins))
            {
                return new List<PinGroup>();
            }

            var spiGroup = new PinGroup(CommonConst.SpiRomPins, "I/O");
            List<Pin> pins = pinMapSheet.GetIoPins();
            foreach (Pin pin in pins)
            {
                if (Regex.IsMatch(pin.PinName, regexPattern, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(pin.PinName, regexPattern1, RegexOptions.IgnoreCase))
                {
                    spiGroup.AddPin(pin.PinName, PinMapConst.TypeIo);
                }
            }
            pinGroups.Add(spiGroup);

            if (spiGroup.PinList.Count != 0)
            {
                var spiTd = new PinGroup(CommonConst.SpiRomTimeDomain, "I/O");
                spiTd.AddPin(CommonConst.SpiRomPins, "I/O");
                pinGroups.Add(spiTd);
            }
            return pinGroups;
        }

        private List<PinGroup> ModifyPinGroups(PinMapSheet pinMap, List<PinGroup> pinGroups)
        {
            var pinGroupNew = new List<PinGroup>();
            foreach (PinGroup group in pinGroups)
            {
                if (!group.PinList.Any())
                {
                    continue;
                }

                Pin first = group.PinList.First();
                Pin pin = pinMap.PinList.FirstOrDefault(a => a.PinName.Equals(first.PinName, StringComparison.OrdinalIgnoreCase));
                if (pin != null && pin.PinType != "")
                {
                    group.PinType = pin.PinType;
                }

                if (Regex.IsMatch(group.PinName, @"_?DIFF\d?$|_DIFF_", RegexOptions.IgnoreCase)
                    && !group.PinName.EndsWith("_Port", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (Pin item in group.PinList)
                    {
                        item.PinType = "Differential";
                    }

                    group.PinType = "Differential";
                    //check pin order => N need to behind to P
                    if (group.PinList.Count == 2)
                    {
                        int pPin = group.PinList.FindLastIndex(p => Regex.IsMatch(p.PinName, "P$", RegexOptions.IgnoreCase));
                        int nPin = group.PinList.FindLastIndex(p => Regex.IsMatch(p.PinName, "N$", RegexOptions.IgnoreCase));
                        if (pPin != -1 && nPin != -1 && pPin > nPin)
                        {
                            ErrorReportManager.AddError(PreActionErrorType.E_RuleViolationPin_01, "", 1, 0, new string[] { group.PinName });
                        }
                    }
                }
                pinGroupNew.Add(group);
            }
            return pinGroupNew;
        }

        private int CheckIoPinGroupSheetFormat(ExcelWorksheet ws)
        {
            for (int row = 1; row <= ws.Dimension.End.Row; row++)
            {
                for (int col = 1; col < ws.Dimension.End.Column; col++)
                {
                    string header = EpplusExtensions.GetCellValue(ws, row, col);
                    if (header.Equals("Pin Group Name", StringComparison.OrdinalIgnoreCase))
                    {
                        return 0;
                    }

                    if (header.Equals("Pin name contained Pin Group (CP)", StringComparison.OrdinalIgnoreCase))
                    {
                        if (EpplusExtensions.GetCellValue(ws, row, col + 1).Equals("Pin name contained Pin Group (FT)", StringComparison.OrdinalIgnoreCase))
                        {
                            return 1;
                        }
                    }
                }
            }
            return -1;
        }

        public class IoPinGroupReaderNew
        {
            public List<string> Header = new List<string> { "Pin name contained Pin Group (CP)", "Pin name contained Pin Group (FT)" };
            public List<PinGroup> PinGroup = new List<PinGroup>();
            private readonly Dictionary<int, string> _pinGroupInd = new Dictionary<int, string>();
            private int _headerIndex;

            public List<PinGroup> ReadMain(ExcelWorksheet sheet)
            {
                ReadHeader(sheet);
                ReadContent(sheet);
                return PinGroup;
            }

            private void ReadHeader(ExcelWorksheet ws)
            {
                for (int i = 1; i <= ws.Dimension.End.Row; i++)
                {
                    if (EpplusExtensions.GetCellValue(ws, i, 1) == Header[0] && EpplusExtensions.GetCellValue(ws, i, 2) == Header[1])
                    {
                        for (int col = 3; col <= ws.Dimension.End.Column; col++)
                        {
                            string groupName = ws.Cells[i, col].Value.ToString();
                            if (groupName != "")
                            {
                                if (!PinGroup.Exists(p => p.PinName.Equals(groupName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    var group = new PinGroup(groupName, "I/O");
                                    PinGroup.Add(group);
                                    _pinGroupInd.Add(col, groupName);
                                }
                            }
                        }
                        _headerIndex = i;
                        break;
                    }
                }
            }

            private void ReadContent(ExcelWorksheet ws)
            {
                for (int i = _headerIndex + 1; i <= ws.Dimension.End.Row; i++)
                {
                    var pins = new List<string> { EpplusExtensions.GetCellValue(ws, i, 1), EpplusExtensions.GetCellValue(ws, i, 2) };
                    foreach (int col in _pinGroupInd.Keys)
                    {
                        if (EpplusExtensions.GetCellValue(ws, i, col).Equals("X", StringComparison.OrdinalIgnoreCase))
                        {
                            PinGroup group = PinGroup.FirstOrDefault(p => p.PinName.Equals(_pinGroupInd[col]));
                            foreach (string pin in pins)
                            {
                                if (pin != "" && group != null && !group.PinList.Exists(p => p.PinName.Equals(pin, StringComparison.OrdinalIgnoreCase)))
                                {
                                    group.AddPin(pin);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
