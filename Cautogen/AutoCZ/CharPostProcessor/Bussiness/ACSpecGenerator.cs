using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.Controller;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class AcSpecGenerator
    {
        private readonly bool _genCharNotUse;
        protected List<string> AcFullCategoryList;
        public AcSpecGenerator(InputParam inputParam)
        {
            _genCharNotUse = inputParam.GenCharNotUse;
        }
        public void Generate(List<CharPlanSheet> charPlanSheets)
        {
            AcSpecSheet preGen = LocalSpecs.TestProgram.AcSpecSheets.FirstOrDefault();
            if (preGen == null)
            {
                LocalSpecs.MessageWriter.WriteLine("[Error] Could not found AC_Spec in the prod progam!");
                return;
            }

            foreach (CharPlanSheet planSheet in charPlanSheets)
            {
                foreach (CharPlanItem charList in planSheet.CharList)
                {
                    foreach (string manual in charList.ManualAc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        ProcessOneManualAc(manual, charList, preGen);
                    }
                }
            }

            WriteAcSpecOutput(preGen);
        }

        private void ProcessOneManualAc(string manual, CharPlanItem charList, AcSpecSheet preGen)
        {
            try
            {
                string manualsymbol = manual.Split(':')[0];
                string shiftSpeed = manual.Split(':')[1];
                string newAcCategory = manual.Split(':')[2];
                string oriAcCategory = manual.Split(':')[3];
                string convertCategory = DataConvertor._ConvertCategoryInAcSpec(newAcCategory, manualsymbol);
                decimal speed = 0L;
                charList.AcCategory = convertCategory;//Replace ac spec to generate.
                string commonCategory = DataConvertor._ConvertCategoryInAcSpec("Common", shiftSpeed).Replace(".", "p");
                if (!string.IsNullOrEmpty(shiftSpeed) && decimal.TryParse(shiftSpeed, out speed))
                {
                    shiftSpeed = (speed * 1000000).ToString();
                }

                AcFullCategoryList = preGen.CategoryList;

                string symbol = ResolveSymbol(preGen, manualsymbol);
                if (string.IsNullOrEmpty(symbol))
                {
                    symbol = manualsymbol;
                    preGen.Rows.Append(WriteOneRecordToAcSpec(symbol, shiftSpeed, convertCategory,
                        shiftSpeed, shiftSpeed, shiftSpeed)); ////
                }

                if (!preGen.CategoryList.Contains(convertCategory) && (charList.Use || _genCharNotUse)) //////
                {
                    AddCategoryWithFollowOrDefault(preGen, convertCategory, oriAcCategory, symbol, shiftSpeed);
                }
                else //manual ac exist in Ac_Spec
                {
                    UpdateCategorySpeed(preGen, symbol, shiftSpeed, convertCategory);
                }
                // if need to generate nwire => reference AC setting need to exist in Timeset_nwire
                if (ProdProg.NwireTimeset != null)
                {
                    HandleNwireTimeset(preGen, charList, symbol, shiftSpeed, oriAcCategory, commonCategory);
                }
            }
            catch (Exception)
            {
            }
        }

        private static string ResolveSymbol(AcSpecSheet preGen, string manualsymbol)
        {
            foreach (AcSpec tmpAc in preGen.Rows)
            {
                if (tmpAc.Symbol.Replace("_", "").Replace("Shmoo", "")
                    .Equals(manualsymbol.Trim(' '), StringComparison.CurrentCultureIgnoreCase))
                {
                    return tmpAc.Symbol;
                }
            }
            return "";
        }

        private void AddCategoryWithFollowOrDefault(AcSpecSheet preGen, string convertCategory, string oriAcCategory, string symbol, string shiftSpeed)
        {
            preGen.CategoryList.Add(convertCategory); //////

            foreach (AcSpec acData in preGen.Rows)
            {
                acData.SelectorList.Add(new Selector("Typ", "Typ"));
                acData.SelectorList.Add(new Selector("Min", "Min"));
                acData.SelectorList.Add(new Selector("Max", "Max"));

                if (acData.CategoryList.Exists(
                        x => x.Name.Equals(oriAcCategory,
                            StringComparison.CurrentCultureIgnoreCase))) //To generate symbol default value follow block ac
                {
                    AddCategoryFollowBlock(acData, convertCategory, oriAcCategory, symbol, shiftSpeed);
                }
                else //To generate symbol default value -1
                {
                    AddCategoryDefaultMinusOne(acData, convertCategory, symbol, shiftSpeed);
                }
            }
        }

        private static void AddCategoryFollowBlock(AcSpec acData, string convertCategory, string oriAcCategory, string symbol, string shiftSpeed)
        {
            if (acData.Symbol.Equals(symbol, StringComparison.CurrentCultureIgnoreCase) &&
                !string.IsNullOrEmpty(shiftSpeed))
            {
                acData.AddCategory(new CategoryInSpec(convertCategory, shiftSpeed,
                    shiftSpeed, shiftSpeed)); ///////
            }
            else
            {
                var sameCategory =
                    acData.CategoryList.Where(
                        x => x.Name.Equals(oriAcCategory,
                            StringComparison.CurrentCultureIgnoreCase)).ToList();
                acData.AddCategory(new CategoryInSpec(convertCategory,
                    sameCategory.First().Typ, sameCategory.First().Typ,
                    sameCategory.First().Typ)); //////
            }
        }

        private static void AddCategoryDefaultMinusOne(AcSpec acData, string convertCategory, string symbol, string shiftSpeed)
        {
            if (acData.Symbol.Equals(symbol, StringComparison.CurrentCultureIgnoreCase) &&
                !string.IsNullOrEmpty(shiftSpeed))
            {
                acData.AddCategory(new CategoryInSpec(convertCategory, shiftSpeed,
                    shiftSpeed, shiftSpeed)); //////
            }
            else
            {
                acData.AddCategory(new CategoryInSpec(convertCategory, "-1", "-1", "-1"));
                //////
            }
        }

        private static void UpdateCategorySpeed(AcSpecSheet preGen, string symbol, string shiftSpeed, string convertCategory)
        {
            foreach (AcSpec acData in preGen.Rows)
            {
                if (acData.Symbol.Equals(symbol, StringComparison.CurrentCultureIgnoreCase) &&
                    !string.IsNullOrEmpty(shiftSpeed) &&
                    !acData.CategoryList.Exists(
                        x => x.Max == shiftSpeed && x.Name.Equals(convertCategory))) ///////
                {
                    foreach (CategoryInSpec category in
                            acData.CategoryList.Where(x => x.Name.Equals(convertCategory))) //////
                    {
                        category.Max = shiftSpeed;
                        category.Min = shiftSpeed;
                        category.Typ = shiftSpeed;
                    }
                }
            }
        }

        private void HandleNwireTimeset(AcSpecSheet preGen, CharPlanItem charList, string symbol, string shiftSpeed, string oriAcCategory, string commonCategory)
        {
            var periods = ProdProg.NwireTimeset.Rows.Select(p => p.CyclePeriod).Distinct().ToList();
            if (!periods.Exists(p => Regex.IsMatch(p, symbol, RegexOptions.IgnoreCase)))
            {
                return;
            }

            if (!preGen.CategoryList.Contains(commonCategory) && (charList.Use || _genCharNotUse)) //////
            {
                AddCommonCategoryWithFollowOrDefault(preGen, commonCategory, oriAcCategory, symbol, shiftSpeed);
            }
            else //manual ac exist in Ac_Spec
            {
                UpdateCategorySpeed(preGen, symbol, shiftSpeed, commonCategory);
            }
        }

        private static void AddCommonCategoryWithFollowOrDefault(AcSpecSheet preGen, string commonCategory, string oriAcCategory, string symbol, string shiftSpeed)
        {
            preGen.CategoryList.Add(commonCategory); //////

            foreach (AcSpec acData in preGen.Rows)
            {
                acData.SelectorList.Add(new Selector("Typ", "Typ"));
                acData.SelectorList.Add(new Selector("Min", "Min"));
                acData.SelectorList.Add(new Selector("Max", "Max"));

                if (acData.CategoryList.Exists(
                        x => x.Name.Equals("Common",
                            StringComparison.CurrentCultureIgnoreCase))) //To generate symbol default value follow block ac
                {
                    AddCommonCategoryFollowBlock(acData, commonCategory, oriAcCategory, symbol, shiftSpeed);
                }
                else //To generate symbol default value -1
                {
                    AddCategoryDefaultMinusOne(acData, commonCategory, symbol, shiftSpeed);
                }
            }
        }

        private static void AddCommonCategoryFollowBlock(AcSpec acData, string commonCategory, string oriAcCategory, string symbol, string shiftSpeed)
        {
            if (acData.Symbol.Equals(symbol, StringComparison.CurrentCultureIgnoreCase) &&
                !string.IsNullOrEmpty(shiftSpeed))
            {
                acData.AddCategory(new CategoryInSpec(commonCategory, shiftSpeed,
                    shiftSpeed, shiftSpeed)); ///////
                return;
            }

            var sameCategory =
                acData.CategoryList.Where(
                    x => x.Name.Equals(oriAcCategory,
                        StringComparison.CurrentCultureIgnoreCase)).ToList();
            if (!sameCategory.Any())
            {
                return;
            }

            acData.AddCategory(new CategoryInSpec(commonCategory,
                sameCategory.First().Typ, sameCategory.First().Typ,
                sameCategory.First().Typ)); //////
        }

        private static void WriteAcSpecOutput(AcSpecSheet preGen)
        {
            // export bin table whose path is decided the GenTexOnly flag
            string outputFolder = Path.Combine(LocalSpecs.OutputFolder, ConstData.AcFolder);
            if (LocalSpecs.InputParam.GenTxtOnly)
            {
                outputFolder = LocalSpecs.OutputFolder;
            }

            string mainAcFile = Path.Combine(outputFolder, "AC_Specs" + ".txt");
            preGen.Write(mainAcFile);
            LocalSpecs.GenSheets.Add(preGen);
        }

        public static void GenerateByPlanTimeset(List<CharPlanSheet> charPlanSheets)
        {
            #region read optional AC from charplan timesettings sheet
            Dictionary<string, Dictionary<string, string>> optionalAc = ReadManualAcSheet();
            if (optionalAc == null)
            {
                return;
            }
            #endregion

            AcSpecSheet preGen = LocalSpecs.TestProgram.AcSpecSheets.FirstOrDefault();
            if (preGen == null)
            {
                LocalSpecs.MessageWriter.WriteLine("[Error] Could not found AC_Spec in the prod progam!");
                return;
            }

            foreach (CharPlanSheet planSheet in charPlanSheets)
            {
                foreach (CharPlanItem charRow in planSheet.CharList)
                {
                    if (string.IsNullOrEmpty(charRow.ShiftFreq) && string.IsNullOrEmpty(charRow.ManualACfromTimeset))
                    {
                        continue;
                    }

                    charRow.AcCategory = charRow.AcCategoryOri;//Replace if ac spec be generated by pin.
                    string shiftSpeed = charRow.ShiftFreq;
                    string acSpecOnOptionalSheet = charRow.ManualACfromTimeset;
                    string blockAc = charRow.AcCategory.Replace("_" + charRow.ShiftFreq + "MHz", "");
                    string shiftSymbol = "ShiftIn_Freq_VAR";

                    if (!string.IsNullOrEmpty(charRow.ShiftFreq) && int.TryParse(charRow.ShiftFreq, out int speed))
                    {
                        shiftSpeed = speed < 1000 ? shiftSpeed = (speed * 1000000).ToString() : shiftSpeed;

                        if (!preGen.CategoryList.Contains(charRow.AcCategory) && charRow.Use)
                        {
                            preGen.CategoryList.Add(charRow.AcCategory);
                            foreach (AcSpec acData in preGen.Rows)
                            {
                                if (
                                    acData.CategoryList.Exists(
                                        x => x.Name.Equals(blockAc, StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    acData.SelectorList.Add(new Selector("Typ", "Typ"));
                                    acData.SelectorList.Add(new Selector("Min", "Min"));
                                    acData.SelectorList.Add(new Selector("Max", "Max"));

                                    if (acData.Symbol.Equals(shiftSymbol, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        acData.AddCategory(new CategoryInSpec(charRow.AcCategory, shiftSpeed,
                                            shiftSpeed, shiftSpeed));
                                    }
                                    else
                                    {
                                        var blockCategory =
                                            acData.CategoryList.Where(
                                                x => x.Name.Equals(blockAc, StringComparison.CurrentCultureIgnoreCase))
                                                .ToList();

                                        acData.AddCategory(new CategoryInSpec(charRow.AcCategory,
                                            blockCategory.First().Typ, blockCategory.First().Typ,
                                            blockCategory.First().Typ));
                                    }
                                }
                            }
                        }
                        blockAc = charRow.AcCategory;
                    }
                    if (!string.IsNullOrEmpty(charRow.ManualACfromTimeset) &&
                        optionalAc.Keys.Any(
                            x => string.Equals(x, charRow.ManualACfromTimeset, StringComparison.OrdinalIgnoreCase)))
                    {
                        charRow.AcCategory = charRow.AcCategoryOri + "_" + charRow.ManualACfromTimeset;
                        Dictionary<string, string> optional =
                            optionalAc.FirstOrDefault(
                                x =>
                                    string.Equals(x.Key, charRow.ManualACfromTimeset, StringComparison.OrdinalIgnoreCase)).Value;
                        string acToGen = charRow.AcCategory;
                        if (!preGen.CategoryList.Contains(acToGen) && charRow.Use)
                        {
                            preGen.CategoryList.Add(acToGen);
                            foreach (AcSpec acData in preGen.Rows)
                            {
                                if (
                                    acData.CategoryList.Exists(
                                        x => x.Name.Equals(blockAc, StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    acData.SelectorList.Add(new Selector("Typ", "Typ"));
                                    acData.SelectorList.Add(new Selector("Min", "Min"));
                                    acData.SelectorList.Add(new Selector("Max", "Max"));

                                    KeyValuePair<string, string> optionalSymbol =
                                        optional.FirstOrDefault(
                                            x => string.Equals(x.Key, acData.Symbol, StringComparison.OrdinalIgnoreCase));
                                    if (!string.IsNullOrEmpty(optionalSymbol.Key) && !string.IsNullOrEmpty(optionalSymbol.Value))
                                    {
                                        acData.AddCategory(new CategoryInSpec(acToGen, optionalSymbol.Value,
                                            optionalSymbol.Value, optionalSymbol.Value));
                                    }
                                    else
                                    {
                                        var blockCategory =
                                            acData.CategoryList.Where(
                                                x => x.Name.Equals(blockAc, StringComparison.CurrentCultureIgnoreCase))
                                                .ToList();

                                        acData.AddCategory(new CategoryInSpec(acToGen,
                                            blockCategory.First().Typ, blockCategory.First().Typ,
                                            blockCategory.First().Typ));
                                    }
                                }
                            }
                        }
                    }
                }
            }
            string outputFolder = Path.Combine(LocalSpecs.OutputFolder, ConstData.AcFolder);
            if (LocalSpecs.InputParam.GenTxtOnly)
            {
                outputFolder = LocalSpecs.OutputFolder;
            }

            string mainAcFile = Path.Combine(outputFolder, "AC_Specs" + ".txt");
            preGen.Write(mainAcFile);
        }

        public void GenerateByTimeSettingsSheet(List<CharPlanSheet> charPlanSheets)
        {
            Dictionary<string, Dictionary<string, string>> manualAc = ReadManualAcSheet();
            if (manualAc == null)
            {
                return;
            }

            AcSpecSheet preGen = LocalSpecs.TestProgram.AcSpecSheets.FirstOrDefault();
            if (preGen == null)
            {
                LocalSpecs.MessageWriter.WriteLine("[Error] Could not found AC_Spec in the prod progam!");
                return;
            }

            foreach (KeyValuePair<string, Dictionary<string, string>> manualCate in manualAc)
            {
                if (preGen.CategoryList.Contains(manualCate.Key))
                {
                    UpdateExistingCategoryFromManual(preGen, manualCate);
                }
                else
                {
                    AddNewManualCategory(preGen, manualCate);
                }
            }
            string outputFolder = Path.Combine(LocalSpecs.OutputFolder, ConstData.AcFolder);
            if (LocalSpecs.InputParam.GenTxtOnly)
            {
                outputFolder = LocalSpecs.OutputFolder;
            }

            string mainAcFile = Path.Combine(outputFolder, "AC_Specs" + ".txt");
            preGen.Write(mainAcFile);
            LocalSpecs.GenSheets.Add(preGen);
            foreach (CharPlanSheet planSheet in charPlanSheets)
            {
                foreach (CharPlanItem charRow in planSheet.CharList)
                {
                    if (string.IsNullOrEmpty(charRow.ManualACfromTimeset))
                    {
                        continue;
                    }

                    charRow.AcCategory = charRow.AcCategoryOri; //Replace if ac spec be generated by pin.
                }
            }
        }

        private static void UpdateExistingCategoryFromManual(AcSpecSheet preGen, KeyValuePair<string, Dictionary<string, string>> manualCate)
        {
            foreach (KeyValuePair<string, string> manualSymbol in manualCate.Value)
            {
                foreach (AcSpec acRow in preGen.Rows)
                {
                    if (acRow.Symbol.Equals(manualSymbol.Key, StringComparison.CurrentCultureIgnoreCase) &&
                        !string.IsNullOrEmpty(manualSymbol.Value))
                    {
                        foreach (CategoryInSpec category in acRow.CategoryList.Where(x => x.Name.Equals(manualCate.Key)))
                        {
                            category.Max = manualSymbol.Value;
                            category.Min = manualSymbol.Value;
                            category.Typ = manualSymbol.Value;
                        }
                    }
                }
            }
        }

        private static void AddNewManualCategory(AcSpecSheet preGen, KeyValuePair<string, Dictionary<string, string>> manualCate)
        {
            preGen.CategoryList.Add(manualCate.Key);
            string blockAc = GetManualBlockAc(manualCate.Key, preGen.CategoryList);
            string shiftFreq = "";
            if (!Regex.IsMatch(blockAc, @"\d+MHz", RegexOptions.IgnoreCase))
            {
                shiftFreq = Regex.Match(manualCate.Key.Split('_').Last(), @"\d+MHz", RegexOptions.IgnoreCase).Value;
                shiftFreq = Regex.Match(shiftFreq, @"\d+").Value;
                if (!string.IsNullOrEmpty(shiftFreq) && int.TryParse(shiftFreq, out int speed))
                {
                    if (speed < 1000)
                    {
                        shiftFreq = (speed * 1000000).ToString();
                    }
                }
            }
            foreach (AcSpec acRow in preGen.Rows)
            {
                AddRowCategoryFromManual(acRow, manualCate, blockAc, shiftFreq);
            }
        }

        private static void AddRowCategoryFromManual(AcSpec acRow, KeyValuePair<string, Dictionary<string, string>> manualCate, string blockAc, string shiftFreq)
        {
            acRow.SelectorList.Add(new Selector("Typ", "Typ"));
            acRow.SelectorList.Add(new Selector("Min", "Min"));
            acRow.SelectorList.Add(new Selector("Max", "Max"));

            KeyValuePair<string, string> manualSymbol = manualCate.Value.FirstOrDefault(x => string.Equals(x.Key, acRow.Symbol, StringComparison.CurrentCultureIgnoreCase));

            if (manualSymbol.Key != null && !string.IsNullOrEmpty(manualSymbol.Value))
            {
                acRow.AddCategory(new CategoryInSpec(
                    manualCate.Key,
                    manualSymbol.Value,
                    manualSymbol.Value,
                    manualSymbol.Value));
            }
            else if (acRow.Symbol.Equals("SHIFT_FREQ", StringComparison.CurrentCultureIgnoreCase) &&
                !string.IsNullOrEmpty(shiftFreq))
            {
                acRow.AddCategory(new CategoryInSpec(
                    manualCate.Key,
                    shiftFreq,
                    shiftFreq,
                    shiftFreq));
            }
            else
            {
                if (acRow.CategoryList.Exists(x => x.Name.Equals(blockAc, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var sameCategory = acRow.CategoryList.Where(x => x.Name.Equals(blockAc, StringComparison.CurrentCultureIgnoreCase)).ToList();
                    acRow.AddCategory(new CategoryInSpec(
                        manualCate.Key,
                        sameCategory.First().Typ,
                        sameCategory.First().Typ,
                        sameCategory.First().Typ));
                }
                else
                {
                    acRow.AddCategory(new CategoryInSpec(manualCate.Key, "-1", "-1", "-1"));
                }
            }
        }

        private static string GetManualBlockAc(string manualCate, List<string> existedCates)
        {
            string result = "";
            string manualFreq = Regex.Match(manualCate, @"\d+Mhz").Value;
            foreach (string cate in existedCates)
            {
                if (!string.IsNullOrEmpty(Regex.Match(manualCate, cate, RegexOptions.IgnoreCase).Value))
                {
                    result = Regex.Match(manualCate, cate, RegexOptions.IgnoreCase).Value;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(result) && !string.IsNullOrEmpty(manualFreq) &&
                existedCates.Any(x => string.Equals(x, result + "_" + manualFreq, StringComparison.OrdinalIgnoreCase)))
            {
                result = existedCates.FirstOrDefault(x => string.Equals(x, result + "_" + manualFreq, StringComparison.OrdinalIgnoreCase));
            }

            return result;
        }

        private static Dictionary<string, Dictionary<string, string>> ReadManualAcSheet()
        {
            OfficeOpenXml.ExcelWorksheet manualSheet = LocalSpecs.OptionalTimesettings;
            if (manualSheet == null)
            {
                return null;
            }

            var manualAc = new Dictionary<string, Dictionary<string, string>>();
            for (int colindex = 2; colindex <= manualSheet.Dimension.Columns; colindex++)
            {
                string key = "";
                for (int rowindex = 1; rowindex <= manualSheet.Dimension.Rows; rowindex++)
                {
                    if (string.IsNullOrEmpty(manualSheet.Cells[1, colindex].Text))
                    {
                        continue;
                    }

                    if (rowindex == 1)
                    {
                        if (!string.IsNullOrEmpty(manualSheet.Cells[rowindex, colindex].Value.ToString().Trim()))
                        {
                            manualAc.Add(manualSheet.Cells[rowindex, colindex].Value.ToString().Trim(),
                                new Dictionary<string, string>());
                            key = manualSheet.Cells[rowindex, colindex].Value.ToString().Trim();
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(manualSheet.Cells[rowindex, colindex].Text))
                        {
                            if (manualAc[key].ContainsKey(manualSheet.Cells[rowindex, 1].Value.ToString().Trim()))
                            {
                                manualAc[key][manualSheet.Cells[rowindex, 1].Value.ToString().Trim()] =
                                    manualSheet.Cells[rowindex, colindex].Value.ToString().Trim();
                            }
                            else
                            {
                                manualAc[key].Add(manualSheet.Cells[rowindex, 1].Value.ToString().Trim(),
                                    manualSheet.Cells[rowindex, colindex].Value.ToString().Trim());
                            }
                        }
                    }
                }
            }
            return manualAc;
        }

        private AcSpec WriteOneRecordToAcSpec(string pStrSymbol, string pStrValue, string pStrCategory, string pStrTyp, string pStrMin, string pStrMax)
        {
            //Write basic data
            var acSpecs = new AcSpec(pStrSymbol, GetSelectorList(), pStrValue);
            //Write Category
            foreach (string categroyName in AcFullCategoryList)
            {
                var categroy = new CategoryInSpec(categroyName, pStrTyp, pStrMin, pStrMax);
                if (!categroyName.Equals(pStrCategory))
                {
                    categroy = new CategoryInSpec(categroyName, "-1", "-1", "-1");
                }
                acSpecs.AddCategory(categroy);
            }
            return acSpecs;
        }

        private static List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }



    }
}
