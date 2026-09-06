using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows
{
    public class IoLevelsGenerator : LevelsGenerator
    {
        public override void UpdateLevelSheet(ref LevelSheet levelSheet, LevelData levelData)
        {
            foreach (IoInfoRow ioInfoRow in levelData.IoInfoRows.Where(row => !row.IsCustom))
            {
                string dcParameterSyntax = !ioInfoRow.IsMerged ? levelData.DcParameterSyntax : "";
                string glbSymbolSuffix = !ioInfoRow.IsMerged ? levelData.GlbSymbolSuffix : "";

                string pinName = ioInfoRow.PinGrpName;
                string vil = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "Vil", dcParameterSyntax, levelData.SplitDcBlockSyntax);
                string vih = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "Vih", dcParameterSyntax, levelData.SplitDcBlockSyntax);
                string vol = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "Vol", dcParameterSyntax, levelData.SplitDcBlockSyntax);
                string voh = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "Voh", dcParameterSyntax, levelData.SplitDcBlockSyntax);
                string vt = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "Vt", dcParameterSyntax, levelData.SplitDcBlockSyntax);
                string vohAlt1 = "0";
                string vohAtl2 = "0";
                string voutLoTyp = "0";
                string voutHiTyp = "0";
                string iol = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Iol", glbSymbolSuffix);
                string ioh = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Ioh", glbSymbolSuffix);
                string vcl = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Vcl", glbSymbolSuffix);
                string vch = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Vch", glbSymbolSuffix);
                string driverMode = ioInfoRow.DriverMode;
                levelSheet.AddIoPinLevel(new IoLevel(pinName, vil, vih, vol, voh, vohAlt1, vohAtl2, iol, ioh, vt, vcl, vch, voutLoTyp, voutHiTyp, driverMode));
            }
        }
    }
}
