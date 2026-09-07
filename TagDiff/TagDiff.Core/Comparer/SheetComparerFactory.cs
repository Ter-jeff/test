using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace TagDiff.Core.Comparer
{
    public class SheetComparerFactory(bool onlyCompareCsharpInstance)
    {
        private readonly bool _onlyCompareCsharpInstance = onlyCompareCsharpInstance;

        public SheetComparerBase? GetIgxlSheetComparer(string sheetName, EnumSheetType enumSheetType, List<SubFlowSheet> subFlowSheets, List<InstanceRow> instanceRows, string job, string baseFile, string compareFile)
        {
            if (sheetName.StartsWithIgnoreCase("Reg_Assign"))
            {
                return new RegSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTTestInstancesSheet)
            {
                return new TestInstanceSheetComparer(sheetName, baseFile, compareFile, subFlowSheets, _onlyCompareCsharpInstance);
            }

            if (enumSheetType == EnumSheetType.DTFlowtableSheet)
            {
                return new FlowSheetComparer(sheetName, baseFile, compareFile, job);
            }

            if (enumSheetType == EnumSheetType.DTACSpecSheet)
            {
                return new AcSpecsSheetComparer(sheetName, baseFile, compareFile, instanceRows);
            }

            if (enumSheetType == EnumSheetType.DTDCSpecSheet)
            {
                return new DcSpecsSheetComparer(sheetName, baseFile, compareFile, instanceRows);
            }

            if (enumSheetType == EnumSheetType.DTBintablesSheet)
            {
                return new BinTableSheetComparer(sheetName, baseFile, compareFile, subFlowSheets);
            }

            if (enumSheetType == EnumSheetType.DTPinMap)
            {
                return new PinMapSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTCharacterizationSheet)
            {
                return new CharacterizationSheetComparer(sheetName, baseFile, compareFile, subFlowSheets);
            }

            if (enumSheetType == EnumSheetType.DTGlobalSpecSheet)
            {
                return new GlobalSpecsSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTLevelSheet)
            {
                var useList = new HashSet<string>([.. instanceRows.Select(x => x.PinLevels).Distinct()], StringExtensions.IgnoreCase);
                if (useList.Contains(sheetName))
                {
                    return new LevelSheetComparer(sheetName, baseFile, compareFile);
                }

                return null;
            }

            if (enumSheetType == EnumSheetType.DTTimesetBasicSheet)
            {
                var useList = new HashSet<string>([.. instanceRows.Select(x => x.TimeSets).Distinct()], StringExtensions.IgnoreCase);
                if (useList.Contains(sheetName))
                {
                    return new TimeSetsSheetComparer(sheetName, baseFile, compareFile);
                }

                return null;
            }

            if (enumSheetType == EnumSheetType.DTPortMapSheet)
            {
                return new PortMapSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTPatternSetSheet)
            {
                return new PatternSetSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTJobListSheet)
            {
                return new JobLisSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTChanMap)
            {
                return new ChanMapSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTPatternSubroutineSheet)
            {
                return new PatternSubroutineSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.DTReferencesSheet)
            {
                return new ReferencesSheetComparer(sheetName, baseFile, compareFile);
            }

            if (enumSheetType == EnumSheetType.BinCut_eqn)
            {
                return new BinCutEqnComparer(sheetName, baseFile, compareFile);
            }

            return new OtherSheetComparer(sheetName, baseFile, compareFile);
        }
    }
}
