using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows
{
    public abstract class LevelsGenerator
    {
        public abstract void UpdateLevelSheet(ref LevelSheet levelSheet, LevelData levelData = null);
    }
}
