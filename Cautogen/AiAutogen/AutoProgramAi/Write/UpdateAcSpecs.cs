using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateAcSpecs
    {
        private const string AcSpecDefault = "-1";

        public AcSpecSheet Work(AcSpecSheet acSpecSheet, List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            AcSpecSheetUpdateEquation(acSpecSheet, comTimeSetBasicSheets);

            //AddComment(acSpecSheet);

            return acSpecSheet;
        }

        private void AcSpecSheetUpdateEquation(AcSpecSheet acSpecSheet,
            List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            if (!comTimeSetBasicSheets.Any())
            {
                return;
            }

            foreach (var timeSetBasicSheet in comTimeSetBasicSheets)
            {
                var dic = new Dictionary<string, double>();
                foreach (var tsetEqnVarMap in timeSetBasicSheet.AllTSetEqnVariable)
                    foreach (var variable in tsetEqnVarMap.DictVariable)
                    {
                        if (!dic.ContainsKey(variable.Key.ToUpper()))
                        {
                            dic.Add(variable.Key.ToUpper(), variable.Value);
                        }
                    }

                //var blockTypeName = GetTimeSetBlockType(timeSetBasicSheet.SheetName);
                ////TimeSet Sheet default block name, ex: SocMbist/GfxScan
                //var categoryName = GetTimeSetCategory(timeSetBasicSheet.SheetName);
                var tokens = timeSetBasicSheet.Name.Split(new[] { '_' },
                    StringSplitOptions.RemoveEmptyEntries).ToList();
                var categoryName = "AC_" + string.Join("_", tokens.GetRange(1, tokens.Count - 1));

                if (dic.Count > 0)
                {
                    UpdateSymbolSpecs(acSpecSheet, dic);
                }

                if (!acSpecSheet.CategoryList.Exists(x =>
                        x.Equals(categoryName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    AddAcCategory(acSpecSheet, dic, categoryName, GetBlockCategory(timeSetBasicSheet.Name));
                    if (acSpecSheet.CategoryTimeSetDic.ContainsKey(categoryName))
                    {
                        acSpecSheet.CategoryTimeSetDic[categoryName].Add(timeSetBasicSheet.Name);
                    }
                    else
                    {
                        acSpecSheet.CategoryTimeSetDic.Add(categoryName,
                            new List<string> { timeSetBasicSheet.Name });
                    }
                }
            }
        }

        private string GetBlockCategory(string timesetName)
        {
            var blockCategory = "";
            if (string.IsNullOrEmpty(timesetName) || timesetName.Split('_').Length < 4)
                return blockCategory;
            var domainSegment = timesetName.Split('_')[2];
            var blockSegment = timesetName.Split('_')[3];

            if (Regex.IsMatch(domainSegment, @"^C", RegexOptions.IgnoreCase))
                blockCategory += "Cpu";
            if (Regex.IsMatch(domainSegment, @"^S", RegexOptions.IgnoreCase))
                blockCategory += "Soc";
            if (Regex.IsMatch(domainSegment, @"^L", RegexOptions.IgnoreCase))
                blockCategory += "Gfx";

            if (string.IsNullOrEmpty(blockCategory))
                return blockCategory;

            if (Regex.IsMatch(blockSegment, @"^SC", RegexOptions.IgnoreCase))
                blockCategory += "Scan";
            if (Regex.IsMatch(blockSegment, @"^BI", RegexOptions.IgnoreCase))
                blockCategory += "Mbist";


            return blockCategory;
        }

        private void AddAcCategory(AcSpecSheet acSpecSheet, Dictionary<string, double> varDic, string categoryName
            , string copyCategoryName = "")
        {
            if (acSpecSheet.CategoryList.Contains(categoryName))
            {
                return;
            }

            acSpecSheet.CategoryList.Add(categoryName);
            foreach (var specs in acSpecSheet.Rows)
            {
                var categoryInSpec = specs.CategoryList.First();
                if (!string.IsNullOrEmpty(copyCategoryName))
                {
                    if (specs.CategoryList.Exists(x =>
                            x.Name.Equals(copyCategoryName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        categoryInSpec = specs.CategoryList.Find(x =>
                            x.Name.Equals(copyCategoryName, StringComparison.CurrentCultureIgnoreCase));
                    }
                }

                var value = !varDic.ContainsKey(specs.Symbol.ToUpper())
                    ? categoryInSpec.Typ
                    : varDic[specs.Symbol.ToUpper()].ToString(CultureInfo.InvariantCulture);
                var newCategory = new CategoryInSpec(categoryName, value, value, value);
                specs.AddCategory(newCategory);
            }
        }

        private void UpdateSymbolSpecs(AcSpecSheet acSpecSheet, Dictionary<string, double> tSetVarValueDict)
        {
            foreach (var timing in tSetVarValueDict)
            {
                if (acSpecSheet.Rows.Exists(x =>
                        x.Symbol.Equals(timing.Key, StringComparison.CurrentCultureIgnoreCase)))
                {
                    continue;
                }

                acSpecSheet.AddAcSpecs(timing.Key, AcSpecDefault, AcSpecDefault, AcSpecDefault, AcSpecDefault);
            }
        }
    }
}
