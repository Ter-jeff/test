using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows
{
    public class CustomIoLevelsGenerator : LevelsGenerator
    {
        private readonly PinMapSheet _pinMapSheet;
        public CustomIoLevelsGenerator(PinMapSheet pinMapSheet)
        {
            _pinMapSheet = pinMapSheet;
        }

        public override void UpdateLevelSheet(ref LevelSheet levelSheet, LevelData levelData)
        {
            LevelSheet tempCustomLevelSheet = new LevelSheet("Levels_TempCustom");
            HashSet<string> customLevelPins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (IoInfoRow ioInfoRow in levelData.IoInfoRows.Where(row => row.IsCustom))
            {
                string pinName = ioInfoRow.PinGrpName;
                string vil = GetDcValueOrDefault(ioInfoRow.Vil, pinName, "Vil", levelData.SplitDcBlockSyntax, levelData.CustomInheritance);
                string vih = GetDcValueOrDefault(ioInfoRow.Vih, pinName, "Vih", levelData.SplitDcBlockSyntax, levelData.CustomInheritance);
                string vol = GetDcValueOrDefault(ioInfoRow.Vol, pinName, "Vol", levelData.SplitDcBlockSyntax, levelData.CustomInheritance);
                string voh = GetDcValueOrDefault(ioInfoRow.Voh, pinName, "Voh", levelData.SplitDcBlockSyntax, levelData.CustomInheritance);
                string vt = GetDcValueOrDefault(ioInfoRow.Vt, pinName, "Vt", levelData.SplitDcBlockSyntax, levelData.CustomInheritance);
                string vohAlt1 = "0";
                string vohAtl2 = "0";
                string voutLoTyp = "0";
                string voutHiTyp = "0";
                string iol = GetGlbValueOrDefault(ioInfoRow.Iol, pinName, "Iol", levelData.GlbSymbolSuffix, levelData.CustomInheritance);
                string ioh = GetGlbValueOrDefault(ioInfoRow.Ioh, pinName, "Ioh", levelData.GlbSymbolSuffix, levelData.CustomInheritance);
                string vcl = GetGlbValueOrDefault(ioInfoRow.Vcl, pinName, "Vcl", levelData.GlbSymbolSuffix, levelData.CustomInheritance);
                string vch = GetGlbValueOrDefault(ioInfoRow.Vch, pinName, "Vch", levelData.GlbSymbolSuffix, levelData.CustomInheritance);
                string driverMode = ioInfoRow.DriverMode;
                customLevelPins.AddRange(_pinMapSheet.DecompGroups(pinName));
                tempCustomLevelSheet.AddIoPinLevel(new IoLevel(pinName, vil, vih, vol, voh, vohAlt1, vohAtl2, iol, ioh, vt, vcl, vch,
                        voutLoTyp, voutHiTyp, driverMode));
            }
            RemoveDuplicatePinLevels(customLevelPins, ref levelSheet, levelData.LevelSheetName);
            levelSheet.Rows.AddRange(tempCustomLevelSheet.Rows);
        }

        private string GetDcValueOrDefault(string infoValue, string pinName, string type, string splitDcBlockSyntax, bool customInheritance)
        {
            if (string.IsNullOrEmpty(infoValue) && customInheritance)
            {
                return SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, type, "", splitDcBlockSyntax);
            }
            return infoValue;
        }

        private string GetGlbValueOrDefault(string infoValue, string pinName, string type, string glbSymbolSuffix, bool customInheritance)
        {
            if (string.IsNullOrEmpty(infoValue) && customInheritance)
            {
                return SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, type, glbSymbolSuffix);
            }
            return infoValue;
        }

        private void RemoveDuplicatePinLevels(HashSet<string> customLevelPins, ref LevelSheet levelSheet, string levelSheetName)
        {
            if (!customLevelPins.Any())
            {
                return;
            }

            HashSet<string> existLevelPins = levelSheet.Rows.Select(x => x.PinName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (string existLevelPin in existLevelPins)
            {
                HashSet<string> existPinGroupPins = _pinMapSheet.DecompGroups(existLevelPin).ToHashSet(StringComparer.OrdinalIgnoreCase);
                if (existPinGroupPins.IsSubsetOf(customLevelPins))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_DuplicateItems_02, "", 1, 0,
                        $"IoInfo sheet contains duplicate pin level: {existLevelPin} definition duplicate remove from {levelSheetName}"
                        , new string[] { existLevelPin, levelSheetName });
                    levelSheet.Rows.RemoveAll(x => x.PinName.Equals(existLevelPin, StringComparison.OrdinalIgnoreCase));
                }
            }
        }
    }
}
