using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.Utility;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class PinMapSheet : IIgxlSheet
    {
        public string SheetType => "DTPinMap";
        public string IgxlSheetName => IgxlSheetNames.PinMap;

        protected string? FilePath;
        protected string? IgxlSheetContext;
        public SourceInfo SourceInfo = new();
        public string Name { get; set; }

        private readonly List<Pin> _pinList = [];
        private readonly List<PinGroup> _groupList = [];
        private readonly Dictionary<string, Pin> _pinDic = new(StringExtensions.IgnoreCase);
        private readonly Dictionary<string, PinGroup> _groupDic = new(StringExtensions.IgnoreCase);
        public IReadOnlyList<Pin> PinList => _pinList;
        public IReadOnlyList<PinGroup> GroupList => _groupList;

        public PinMapSheet(ExcelWorksheet excelWorksheet)
        {
            IgxlSheetContext = excelWorksheet.Cells[1, 1].Text;
            Name = excelWorksheet.Name;
        }

        public PinMapSheet(string sheetName)
        {
            Name = sheetName;
        }

        public void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
            {
                version = "2.1";
            }

            if (!PinList.Any() && !GroupList.Any())
            {
                return;
            }

            string? dir = Path.GetDirectoryName(fileName);
            if (dir != null && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            Dictionary<string, Dictionary<string, SheetInfo>> sheetClassMapping = IgxlSheetHelpers.GetIgxlSheetsVersion();
            if (sheetClassMapping.TryGetValue(IgxlSheetName, out Dictionary<string, SheetInfo>? dic))
            {
                if (dic.TryGetValue(version, out SheetInfo? igxlSheetsVersion1))
                {
                    WriteSheet(fileName, version, igxlSheetsVersion1);
                }
                else if (dic.TryGetValue("2.1", out SheetInfo? igxlSheetsVersion))
                {
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        protected void WriteSheet(string fileName, string version, SheetInfo sheetInfo)
        {
            if (_pinList.Count == 0)
            {
                return;
            }

            using var sw = new StreamWriter(fileName, false);
            var sb = new StringBuilder();
            int maxCount = IgxlSheetHelpers.GetMaxCount(sheetInfo);
            int groupNameIndex = IgxlSheetHelpers.GetIndexFrom(sheetInfo, "Group Name");
            int pinNameIndex = IgxlSheetHelpers.GetIndexFrom(sheetInfo, "Pin Name");
            int typeIndex = IgxlSheetHelpers.GetIndexFrom(sheetInfo, "Type");
            int commentIndex = IgxlSheetHelpers.GetIndexFrom(sheetInfo, "Comment");

            #region headers
            int startRow = sheetInfo.Columns.RowCount;
            sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
            for (int i = 1; i < startRow; i++)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];

                IgxlSheetHelpers.SetField(sheetInfo, i, arr);

                IgxlSheetHelpers.SetColumns(sheetInfo, i, arr);

                IgxlSheetHelpers.WriteHeader(arr, sb);
            }
            #endregion

            #region data
            foreach (Pin row in _pinList)
            {
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                if (!string.IsNullOrEmpty(row.PinName))
                {
                    arr[0] = row.ColumnA ?? "";
                    arr[pinNameIndex] = row.PinName.Replace("/", "");
                    arr[typeIndex] = row.PinType ?? "";
                    arr[commentIndex] = row.Comment ?? "";

                }
                else
                {
                    arr = ["\t"];
                }
                sb.AppendLine(string.Join("\t", arr));
            }

            for (int index = 0; index < _groupList.Count; index++)
            {
                PinGroup group = _groupList[index];
                bool isFirst = true;
                foreach (Pin row in group.PinList)
                {
                    string[] arr;
                    if (!string.IsNullOrEmpty(row.PinName))
                    {
                        arr = [.. Enumerable.Repeat("", maxCount)];
                        arr[0] = row.ColumnA ?? "";
                        arr[groupNameIndex] = group.PinName;
                        arr[pinNameIndex] = row.PinName.Replace("/", "");

                        if (string.IsNullOrEmpty(row.PinType) && HasSinglePin(row.PinName))
                        {
                            row.PinType = GetPin(row.PinName)!.PinType;
                        }

                        arr[typeIndex] = (isFirst ? row.PinType : "") ?? "";
                        arr[commentIndex] = row.Comment ?? "";
                    }
                    else
                    {
                        arr = ["\t"];
                    }
                    sb.AppendLine(string.Join("\t", arr));
                    if (isFirst)
                    {
                        isFirst = false;
                    }
                }
            }
            #endregion

            sw.Write(sb.ToString());
        }

        public void AddRow(PinBase pinBase)
        {
            if (pinBase is Pin pin)
            {
                AddPin(pin);
            }
            else
            {
                AddGroup((PinGroup)pinBase);
            }
        }

        public string GetDiffGroupName(string[] pair)
        {
            if (pair.Length == 2)
            {
                PinGroup? group = _groupList.FirstOrDefault(x => x.PinList.Count == 2 &&
                                                                x.PinList.Exists(p => p.PinName!.EqualsIgnoreCase(pair[0])) &&
                                                                x.PinList.Exists(p => p.PinName!.EqualsIgnoreCase(pair[1])));
                if (group != null)
                {
                    return group.PinName ?? "";
                }
            }
            return "";
        }

        public Dictionary<string, string> GetUartPinDic()
        {
            Dictionary<string, string> pinNameList = new Dictionary<string, string>(StringExtensions.IgnoreCase);
            if (IsGroupExist("UART_TX"))
            {
                pinNameList.Add("UART_TX", GetPinsFromGroup("UART_TX").FirstOrDefault()?.PinName!);
            }
            if (IsGroupExist("UART_RX"))
            {
                pinNameList.Add("UART_RX", GetPinsFromGroup("UART_RX").FirstOrDefault()?.PinName!);
            }
            return pinNameList;
        }

        public List<Pin> GetPowerPins()
        {
            List<Pin> pinList = _pinList.FindAll(p => p.PinType!.EqualsIgnoreCase(PinMapConst.TypePower));
            return pinList;
        }

        public List<string> GetIoPins(List<string> patternPins)
        {
            var pins = new List<string>();
            foreach (string patternPin in patternPins)
            {
                Pin? pin = GetPin(patternPin);
                if (pin != null && pin.PinType == PinMapConst.TypeIo)
                {
                    pins.Add(pin.PinName!);
                }
            }
            return pins;
        }

        public List<Pin> GetIoPins()
        {
            List<Pin> pinList = _pinList.FindAll(p => p.PinType!.EqualsIgnoreCase(PinMapConst.TypeIo));
            return pinList;
        }

        public List<Pin> GetIoContinuityPins()
        {
            List<Pin> pinList = GetAllDigitalDisconnectContinuityPins();
            pinList = [.. pinList.Where(x => !(x.PinName!.StartsWithIgnoreCase("REFCLK_") || x.PinName!.EndsWithIgnoreCase("_PA")))];
            return pinList;
        }

        public List<Pin> GetAllDigitalDisconnectContinuityPins()
        {
            List<Pin> pinList = _pinList.FindAll(p => p.ChannelType!.EqualsIgnoreCase(PinMapConst.TypeIo));
            pinList = [.. pinList.Where(x => !((x.PinName!.StartsWithIgnoreCase("VDD") && x.PinName!.ContainsIgnoreCase("SENSE")) || (x.PinName!.StartsWithIgnoreCase("VDD") && x.PinName!.ContainsIgnoreCase("MONITOR")) || (x.PinName!.StartsWithIgnoreCase("VSS") && x.PinName!.ContainsIgnoreCase("SENSE"))
                ))];
            return pinList;
        }

        #region pin
        public void InsertPinAt(int index, Pin pin)
        {
            _pinList.Insert(index, pin);
            _pinDic[pin.PinName!] = pin;
        }

        public void RemovePinAt(int index)
        {
            if (index < 0 || index >= _pinList.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            Pin pin = _pinList[index];
            _pinList.RemoveAt(index);
            _pinDic.Remove(pin.PinName!);
        }

        public void RemoveGroupAt(int index)
        {
            if (index < 0 || index >= _groupList.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            PinGroup group = _groupList[index];

            _groupList.RemoveAt(index);
            _groupDic.Remove(group.PinName!);
        }

        public bool TryGetPin(string name, [MaybeNullWhen(false)] out Pin pin) => _pinDic.TryGetValue(name, out pin);

        public void AddPin(Pin pin)
        {
            if (!HasSinglePin(pin.PinName!))
            {
                _pinList.Add(pin);
                _pinDic.Add(pin.PinName!, pin);
            }
        }

        public void AddPins(List<Pin> pins)
        {
            foreach (Pin pin in pins)
            {
                AddPin(pin);
            }
        }

        public virtual bool IsPinExist(string pinName)
        {
            return HasSinglePin(pinName) || IsGroupExist(pinName);
        }

        public bool HasSinglePin(string pinName)
        {
            return _pinDic.ContainsKey(pinName);
        }

        public Pin? GetPin(string pinName)
        {
            _pinDic.TryGetValue(pinName, out Pin? pin);
            return pin;
        }

        public string GetPinType(string pinName)
        {
            if (_pinDic.TryGetValue(pinName, out Pin? pin))
            {
                return pin.PinType ?? "";
            }

            return "";
        }
        #endregion

        #region pinGroup
        public bool TryGetGroup(string name, [MaybeNullWhen(false)] out PinGroup pinGroup) => _groupDic.TryGetValue(name, out pinGroup);

        public void AddGroup(PinGroup pinGroup)
        {
            if (!IsGroupExist(pinGroup.PinName!))
            {
                if (pinGroup.PinList.Count == 0)
                {
                    throw new Exception($"There isn't any pin in PinGroup:{pinGroup.PinName} . ");
                }
                _groupList.Add(pinGroup);
                _groupDic.Add(pinGroup.PinName!, pinGroup);
            }
        }

        public void InsertGroup(int index, PinGroup pinGroup)
        {
            if (!IsGroupExist(pinGroup.PinName!))
            {
                if (pinGroup.PinList.Count == 0)
                {
                    throw new Exception($"There isn't any pin in PinGroup:{pinGroup.PinName} . ");
                }
                _groupList.Insert(index, pinGroup);
                _groupDic.Add(pinGroup.PinName!, pinGroup);
            }
        }

        public void AddGroups(List<PinGroup> pinGroups)
        {
            foreach (PinGroup pinGroup in pinGroups)
            {
                if (pinGroup.PinList == null || pinGroup.PinList.Count == 0)
                {
                    continue;
                }

                AddRow(pinGroup);
            }
        }

        public bool IsGroupExist(string pinName)
        {
            return _groupDic.ContainsKey(pinName);
        }

        public PinGroup? GetGroup(string groupName)
        {
            _groupDic.TryGetValue(groupName, out PinGroup? group);
            return group;
        }

        public List<Pin> GetPinsFromGroup(string pinName)
        {
            var resultPins = new List<Pin>();
            if (!_groupDic.TryGetValue(pinName, out PinGroup? group))
            {
                return [];
            }
            foreach (Pin pin in group.PinList)
            {
                if (pin.PinName!.EqualsIgnoreCase(pinName))
                {
                    resultPins.Add(pin);
                    continue;
                }
                if (IsGroupExist(pin.PinName))
                {
                    resultPins.AddRange(GetPinsFromGroup(pin.PinName));
                }
                else
                {
                    _pinDic.TryGetValue(pin.PinName, out Pin? singlePin);
                    resultPins.Add(singlePin ?? new Pin(pin.PinName, "", "Does not exist in Pin map sheet"));
                }
            }
            return resultPins;
        }

        public List<string> DecompGroups(string pinGroup)
        {
            var newPinList = new List<string>();
            if (TryGetGroup(pinGroup, out PinGroup? group))
            {
                foreach (Pin pin in group.PinList)
                {
                    if (pin.PinName!.EqualsIgnoreCase(pinGroup))
                    {
                        newPinList.Add(pin.PinName);
                        continue;
                    }
                    newPinList.AddRange(DecompGroups(pin.PinName!));
                }
            }
            else
            {
                newPinList.Add(pinGroup);
            }
            return newPinList;
        }

        public List<PinGroup> GenDcviGroup()
        {
            var pinGroups = new List<PinGroup>();
            var group1 = PinList.Where(a => a.ChannelType!.StartsWith("DCVI") && a.PinType!.ContainsIgnoreCase(PinMapConst.TypePower)).ToList();
            var pinGroup1 = new PinGroup("All_DCVI", PinMapConst.TypePower);
            pinGroup1.AddPins([.. group1]);
            if (pinGroup1.PinList.Count != 0)
            {
                pinGroups.Add(pinGroup1);
            }

            var group2 = PinList.Where(a => a.ChannelType!.StartsWith("DCVI") && a.PinType!.ContainsIgnoreCase("Analog")).ToList();
            var pinGroup2 = new PinGroup("All_DCVI_" + PinMapConst.TypeAnalog, PinMapConst.TypeAnalog);
            pinGroup2.AddPins([.. group2]);
            if (pinGroup2.PinList.Count != 0)
            {
                pinGroups.Add(pinGroup2);
            }

            return pinGroups;
        }

        public void GenDcvsGroup()
        {
            var group = PinList.Where(a => a.ChannelType!.StartsWith("DCVS") && a.PinType!.ContainsIgnoreCase(PinMapConst.TypePower)).ToList();
            var pinGroup = new PinGroup("All_DCVS", PinMapConst.TypePower);
            pinGroup.AddPins([.. group]);
            if (pinGroup.PinList.Count != 0)
            {
                pinGroup.PinList.ForEach(x => x.ColumnA = "Autogen Default");
                AddGroup(pinGroup);
            }
        }
        #endregion

        public List<string> GetAllPins()
        {
            var pins = PinList.Select(x => x.PinName).ToList();
            pins.AddRange(GroupList.Select(x => x.PinName));
            return pins;
        }
    }
}
