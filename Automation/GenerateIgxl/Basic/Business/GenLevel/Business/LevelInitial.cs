using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.Static;

using TestPlanLib.DataStruct;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business
{
    public class LevelInitial
    {
        public const string Scan = "Scan";
        public const string Mbist = "Mbist";
        public string Rtos = ModuleSingleton.Instance().ModuleRtos;
        public const string HardIp = "HardIP";
        public const string Ids = "IDS";
        public const string Efuse = "Efuse";
        public const string Conti = "Conti";
        public const string BinCut = "BinCut";
        public const string WalkingZPosLevelSheetName = "Levels_WalkingZ_Pos";
        public const string WalkingZNegLevelSheetName = "Levels_WalkingZ_Neg";

        private readonly IReadOnlySet<string> _reservedBlocks = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Common",
            "Scan",
            "HardIP",
            "DCTEST_Continuity"
        };
        private readonly IoInfoSheet _ioInfoSheet;
        private readonly IoInfoSheet _ioInfoConcurrentSheet;
        private readonly bool _isSplitDc = false;
        private readonly ReadOnlyDictionary<string, string> _dcParameterSyntaxDict;
        private readonly ReadOnlyDictionary<string, string> _splitDcBlockSyntaxDict;
        private readonly ReadOnlyDictionary<string, string> _glbSymbolSuffixDict;

        public LevelInitial(IoInfoSheet ioInfoSheet, IoInfoSheet ioInfoConcurrentSheet, bool isSplitDc)
        {
            _ioInfoSheet = ioInfoSheet;
            _ioInfoConcurrentSheet = ioInfoConcurrentSheet;
            _isSplitDc = isSplitDc;
            _dcParameterSyntaxDict = new(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Scan"] = "Scan",
            });
            _splitDcBlockSyntaxDict = new(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Scan"] = isSplitDc ? "SC" : ""
            });
            _glbSymbolSuffixDict = new(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Scan"] = "Scan",
                ["HardIP"] = "H",
                ["DCTEST_Continuity"] = "Conti"
            });
        }

        public Dictionary<string, LevelData> InitialLevelDatas()
        {
            List<LevelData> list = new List<LevelData>();
            List<IoInfoRow> commonIoInfoRows = _ioInfoSheet.GetBlockIoInfo("Common");
            commonIoInfoRows.ForEach(row => row.IsMerged = true);
            list.AddRange(GetIoInfoLevelDatas(_ioInfoSheet, commonIoInfoRows, _isSplitDc));

            if (_ioInfoConcurrentSheet != null)
            {
                list.AddRange(GetConcurrentIoInfoLevelDatas(_ioInfoConcurrentSheet));
            }

            Dictionary<string, LevelData> multiLevelDatas = list.ToDictionary(x => x.LevelSheetName, x => x, StringComparer.OrdinalIgnoreCase);

            SetCustomIoInfoLevelDatas(_ioInfoSheet, commonIoInfoRows, ref multiLevelDatas);
            return multiLevelDatas;
        }

        private static (string Block, string RefBolck) GetInfoBlock(string str)
        {
            string[] segments = str.Split(':', 3);

            return segments.Length switch
            {
                1 => (segments[0], string.Empty),
                _ => (segments[0], segments[1])
            };
        }

        private List<LevelData> GetIoInfoLevelDatas(IoInfoSheet ioInfoSheet, List<IoInfoRow> commonIoInfoRows, bool isSplitDc)
        {
            List<LevelData> list = new List<LevelData>();
            if (BlockStatus.GetAutomationBlockStatus(BlockConst.Basic).Down)
            {
                List<IoInfoRow> contiIoInfoRows = ioInfoSheet.GetBlockIoInfo("DCTEST_Continuity");
                if (!contiIoInfoRows.Any())
                {
                    contiIoInfoRows = commonIoInfoRows;
                }
                list.Add(new LevelData($"Levels_{BlockConst.Conti}", contiIoInfoRows)
                {
                    GlbSymbolSuffix = GetSyntaxByBlock("DCTEST_Continuity", _glbSymbolSuffixDict)
                });

                list.Add(new LevelData(WalkingZPosLevelSheetName, commonIoInfoRows));
                list.Add(new LevelData(WalkingZNegLevelSheetName, commonIoInfoRows));

                if (LocalSpecs.IsModuleIncluded("IDS"))
                {
                    list.Add(new LevelData($"Levels_{BlockConst.Ids}", commonIoInfoRows));
                }
            }

            if (BlockStatus.GetAutomationBlockStatus(BlockConst.Evs).Down && LocalSpecs.IsModuleIncluded(BlockConst.Evs))
            {
                list.Add(new LevelData($"Levels_EVS_{BlockConst.Scan}", GetMergedIoInfoRows(commonIoInfoRows, ioInfoSheet.GetBlockIoInfo(BlockConst.Scan)))
                {
                    SplitDcBlockSyntax = GetSyntaxByBlock(BlockConst.Scan, _splitDcBlockSyntaxDict),
                    GlbSymbolSuffix = GetSyntaxByBlock(BlockConst.Scan, _glbSymbolSuffixDict)
                });
                list.Add(new LevelData($"Levels_EVS_{BlockConst.Mbist}", commonIoInfoRows));
            }

            if (BlockStatus.GetAutomationBlockStatus(BlockConst.Scan).Down && LocalSpecs.IsModuleIncluded(BlockConst.Scan))
            {
                list.Add(new LevelData($"Levels_{BlockConst.Scan}", GetMergedIoInfoRows(commonIoInfoRows, ioInfoSheet.GetBlockIoInfo(BlockConst.Scan)))
                {
                    SplitDcBlockSyntax = GetSyntaxByBlock(BlockConst.Scan, _splitDcBlockSyntaxDict),
                    GlbSymbolSuffix = GetSyntaxByBlock(BlockConst.Scan, _glbSymbolSuffixDict)
                });
            }

            if (BlockStatus.GetAutomationBlockStatus(BlockConst.Rtos).Down && LocalSpecs.IsModuleIncluded(BlockConst.Rtos))
            {
                list.Add(new LevelData($"Levels_{ModuleSingleton.Instance().ModuleRtos}", commonIoInfoRows));
            }

            if (BlockStatus.GetAutomationBlockStatus(BlockConst.Mbist).Down && LocalSpecs.IsModuleIncluded(BlockConst.Mbist))
            {
                list.Add(new LevelData($"Levels_{BlockConst.Mbist}", commonIoInfoRows));
                list.Add(new LevelData($"Levels_{BlockConst.Mbist}_DSSC", commonIoInfoRows));
            }

            if (BlockStatus.GetAutomationBlockStatus(BlockConst.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockConst.Efuse))
            {
                list.Add(new LevelData($"Levels_{BlockConst.Efuse}", commonIoInfoRows));
            }

            if ((BlockStatus.GetAutomationBlockStatus(BlockConst.HardIp).Down || BlockStatus.GetAutomationBlockStatus(BlockConst.Lcd).Down)
                && LocalSpecs.IsModuleIncluded(BlockConst.HardIp))
            {
                list.Add(new LevelData($"Levels_{BlockConst.HardIp}", GetMergedIoInfoRows(commonIoInfoRows, ioInfoSheet.GetBlockIoInfo(BlockConst.HardIp)))
                {
                    GlbSymbolSuffix = GetSyntaxByBlock(BlockConst.HardIp, _glbSymbolSuffixDict)
                });
            }

            if (isSplitDc)
            {
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.BinCut).Down && LocalSpecs.IsModuleIncluded(BlockConst.BinCut))
                {
                    list.Add(new LevelData($"Levels_{BlockConst.BinCut}", commonIoInfoRows));
                }
            }
            return list;
        }

        private List<LevelData> GetConcurrentIoInfoLevelDatas(IoInfoSheet ioInfoConcurrentSheet)
        {
            List<LevelData> list = new List<LevelData>();
            IEnumerable<IGrouping<string, IoInfoRow>> commonConcurrentIoInfoRows = ioInfoConcurrentSheet.GetBlockIoInfo("Common").GroupBy(x => x.TimeDomain);
            commonConcurrentIoInfoRows.SelectMany(x => x).ToList().ForEach(row => row.IsMerged = true);
            foreach (IGrouping<string, IoInfoRow> groupIoInfo in ioInfoConcurrentSheet.GetBlockIoInfo(BlockConst.Scan).GroupBy(x => x.TimeDomain))
            {
                string timeDomainSuffix = string.IsNullOrEmpty(groupIoInfo.Key) ? "" : $"_{groupIoInfo.Key}";
                IGrouping<string, IoInfoRow> sameDomainInCommon = commonConcurrentIoInfoRows.FirstOrDefault(x => x.Key == groupIoInfo.Key);
                List<IoInfoRow> ioInfoRows = sameDomainInCommon == null ? groupIoInfo.ToList() : GetMergedIoInfoRows(sameDomainInCommon.ToList(), groupIoInfo.ToList());
                list.Add(new LevelData("Levels_Con_SC" + timeDomainSuffix, ioInfoRows)
                {
                    DcParameterSyntax = GetSyntaxByBlock(BlockConst.Scan, _dcParameterSyntaxDict),
                    SplitDcBlockSyntax = GetSyntaxByBlock(BlockConst.Scan, _splitDcBlockSyntaxDict),
                    GlbSymbolSuffix = GetSyntaxByBlock(BlockConst.Scan, _glbSymbolSuffixDict)
                });
            }
            foreach (IGrouping<string, IoInfoRow> groupIoInfo in ioInfoConcurrentSheet.GetBlockIoInfo(BlockConst.HardIp).GroupBy(x => x.TimeDomain))
            {
                string timeDomainSuffix = string.IsNullOrEmpty(groupIoInfo.Key) ? "" : $"_{groupIoInfo.Key}";
                IGrouping<string, IoInfoRow> sameDomainInCommon = commonConcurrentIoInfoRows.FirstOrDefault(x => x.Key == groupIoInfo.Key);
                List<IoInfoRow> ioInfoRows = sameDomainInCommon == null ? groupIoInfo.ToList() : GetMergedIoInfoRows(sameDomainInCommon.ToList(), groupIoInfo.ToList());
                list.Add(new LevelData("Levels_Con_H" + timeDomainSuffix, ioInfoRows)
                {
                    GlbSymbolSuffix = GetSyntaxByBlock(BlockConst.HardIp, _glbSymbolSuffixDict)
                });
            }
            foreach (IGrouping<string, IoInfoRow> groupIoInfo in ioInfoConcurrentSheet.GetBlockIoInfo("DCTEST_Continuity").GroupBy(x => x.TimeDomain))
            {
                string timeDomainSuffix = string.IsNullOrEmpty(groupIoInfo.Key) ? "" : $"_{groupIoInfo.Key}";
                IGrouping<string, IoInfoRow> sameDomainInCommon = commonConcurrentIoInfoRows.FirstOrDefault(x => x.Key == groupIoInfo.Key);
                List<IoInfoRow> ioInfoRows = sameDomainInCommon == null ? groupIoInfo.ToList() : GetMergedIoInfoRows(sameDomainInCommon.ToList(), groupIoInfo.ToList());
                list.Add(new LevelData("Levels_Con_Conti" + timeDomainSuffix, ioInfoRows)
                {
                    GlbSymbolSuffix = GetSyntaxByBlock("DCTEST_Continuity", _glbSymbolSuffixDict),
                });
            }
            return list;
        }

        private void SetCustomIoInfoLevelDatas(IoInfoSheet ioInfoSheet, List<IoInfoRow> commonIoInfoRows, ref Dictionary<string, LevelData> multiLevelDatas)
        {
            foreach (KeyValuePair<string, List<IoInfoRow>> info in
                ioInfoSheet.BlockIoInfo.Where(x => x.Key.StartsWith("Levels_", StringComparison.OrdinalIgnoreCase)))
            {
                info.Value.ForEach(row => row.IsCustom = true);
                List<IoInfoRow> refIoInfoRows = commonIoInfoRows;
                bool customInheritance = false;

                (string block, string refBlock) = GetInfoBlock(info.Key);
                if (!string.IsNullOrEmpty(refBlock) && _ioInfoSheet.BlockIoInfo.ContainsKey(refBlock))
                {
                    refIoInfoRows = _ioInfoSheet.GetBlockIoInfo(refBlock);
                    if (_reservedBlocks.Contains(refBlock))
                    {
                        customInheritance = true;
                    }
                }

                multiLevelDatas[block] = new LevelData(info.Key, GetMergedIoInfoRows(refIoInfoRows, info.Value))
                {
                    SplitDcBlockSyntax = GetSyntaxByBlock(refBlock, _splitDcBlockSyntaxDict),
                    GlbSymbolSuffix = GetSyntaxByBlock(refBlock, _glbSymbolSuffixDict),
                    CustomInheritance = customInheritance
                };
            }
        }

        public List<IoInfoRow> GetMergedIoInfoRows(List<IoInfoRow> commonInfoRows, List<IoInfoRow> blockInfoRows)
        {
            List<IoInfoRow> result = new List<IoInfoRow>();
            if (blockInfoRows == null || !blockInfoRows.Any())
            {
                return commonInfoRows;
            }

            HashSet<string> blockPinGroups = blockInfoRows.Select(x => x.PinGrpName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (IoInfoRow commonInfoRow in commonInfoRows)
            {
                if (blockPinGroups.Contains(commonInfoRow.PinGrpName))
                {
                    continue;
                }
                result.Add(commonInfoRow);
            }
            result.AddRange(blockInfoRows);
            return result;
        }

        private string GetSyntaxByBlock(string block, ReadOnlyDictionary<string, string> syntaxDict)
        {
            return syntaxDict.TryGetValue(block, out string result) ? result : "";
        }
    }
}
