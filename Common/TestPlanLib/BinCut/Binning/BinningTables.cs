using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.Binning
{
    public partial class BinningTables : List<BinningTable>
    {
        [GeneratedRegex(@"\d+", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public bool IsTitleFail;
        public string Job = "";
        public bool IsEqnBin
        {
            get { return this[0].EqnBinIdx != -1 && this[0].BinXIdsMaxIdx != -1; }
        }

        public BinningTables GetOtherRailTables()
        {
            var binningTables = new BinningTables();
            foreach (BinningTable table in this)
            {
                BinningTable otherRailTable = table.Copy();
                otherRailTable.Rows = [.. table.Rows.Where(x => x.IsOtherRail)];
                binningTables.Add(otherRailTable);
            }
            return binningTables;
        }

        public BinningTables GetBinningRailTables()
        {
            var binningTables = new BinningTables();
            foreach (BinningTable table in this)
            {
                BinningTable binningRail = table.Copy();
                binningRail.Rows = [.. table.Rows.Where(x => !x.IsOtherRail)];
                binningTables.Add(binningRail);
            }
            return binningTables;
        }

        #region check
        public void Check()
        {
            CheckTitle();

            CheckAllowEqual();

            CheckCmGbvalue();

            CheckIDvalue();
        }

        private void CheckTitle()
        {
            if (Count == 0)
            {
                return;
            }

            IEnumerable<string> titleList = this[0].TitleList.Select(x => x.Item1);
            for (int index = 0; index < Count; index++)
            {
                BinningTable sheet = this[index];
                var result1 = titleList.Except(sheet.TitleList.Select(x => x.Item1), StringExtensions.IgnoreCase).Where(x => !string.IsNullOrEmpty(x)).ToList();
                var result2 = sheet.TitleList.Select(x => x.Item1).Except(titleList, StringExtensions.IgnoreCase).Where(x => !string.IsNullOrEmpty(x)).ToList();
                if (result1.Count != 0)
                {
                    string errorMessage = $"The missing column {string.Join(",", result1)} in sheet {sheet.SheetName} !!!";
                    sheet.AddError(BinCutErrorType.E_FormatError_05, sheet.SheetName, sheet.StartRowIdx + 1, 0, $"The missing column {string.Join(",", result1)} in sheet {sheet.SheetName} !!!", [string.Join(",", result1), sheet.SheetName]);
                    IsTitleFail = true;
                }
                if (result2.Count != 0)
                {
                    string errorMessage = $"The missing column {string.Join(",", result2)} in sheet {this[0].SheetName} !!!";
                    sheet.AddError(BinCutErrorType.E_FormatError_05, this[0].SheetName, this[0].StartRowIdx + 1, 0, $"The missing column {string.Join(",", result2)} in sheet {this[0].SheetName} !!!", [string.Join(",", result2), this[0].SheetName]);
                    IsTitleFail = true;
                }
            }
        }

        private void CheckCmGbvalue()
        {
            var pmodeParameters = new Dictionary<int, List<BinningRow>>();
            int tableIdx = 0;
            foreach (BinningTable table in this)
            {
                if (table.AllowEqualIdx == -1)
                {
                    continue;
                }

                pmodeParameters.Add(tableIdx, [.. table.Rows.Where(row => !string.IsNullOrEmpty(row.RowData[table.AllowEqualIdx]))]);
                tableIdx++;
            }
            for (int idx = 0; idx < pmodeParameters.Count; idx++)
            {
                BinningTable table = this[idx];
                var pmodeGroups = pmodeParameters[idx].GroupBy(x => x.RowData[table.ModeIdx]).ToList();
                for (int i = 0; i < pmodeGroups.Count; i++)
                {
                    IGrouping<string, BinningRow> pmodeGroup = pmodeGroups[i];
                    string domain = pmodeGroup.First().RowData[table.DomainIdx];
                    string allowEqualMode = pmodeGroup.First().RowData[table.AllowEqualIdx];
                    var allowEqualRowDatas = table.Rows.Where(x => x.RowData[table.ModeIdx].EqualsIgnoreCase(allowEqualMode)).ToList();
                    // only check same domain case
                    CheckSameDomain(table, pmodeGroup, domain, allowEqualMode, allowEqualRowDatas);
                }
            }
        }

        private static void CheckSameDomain(BinningTable binningTable, IGrouping<string, BinningRow> pmodeGroup, string domain, string allowEqualMode, List<BinningRow> binningRows)
        {
            if (binningRows.Count != 0 && binningRows.Any(x => x.RowData[binningTable.DomainIdx].EqualsIgnoreCase(domain)))
            {
                if (binningRows.Count != pmodeGroup.Count())
                {
                    // differnet equation count 
                    string errorMessage = $"{allowEqualMode} and {pmodeGroup.First().RowData[binningTable.ModeIdx]} have different number of equation";
                    binningTable.AddError(BinCutErrorType.E_Equation_01, binningTable.SheetName, 0, 0, $"{allowEqualMode} and {pmodeGroup.First().RowData[binningTable.ModeIdx]} have different number of equation", [allowEqualMode, pmodeGroup.First().RowData[binningTable.ModeIdx]]);
                }
                foreach (BinningRow row in pmodeGroup)
                {
                    string eq = binningTable.EqnIdx != -1 ? row.RowData[binningTable.EqnIdx] : "";
                    string c = binningTable.CIdx != -1 ? row.RowData[binningTable.CIdx] : "";
                    string m = binningTable.MIdx != -1 ? row.RowData[binningTable.MIdx] : "";
                    string cpgb = binningTable.CpGbIdx != -1 ? row.RowData[binningTable.CpGbIdx] : "";
                    string cp2Gb = binningTable.Cp2GbIdx != -1 ? row.RowData[binningTable.Cp2GbIdx] : "";
                    string ft1Gb = binningTable.Ft1GbIdx != -1 ? row.RowData[binningTable.Ft1GbIdx] : "";
                    string ft2Gb = binningTable.Ft2GbIdx != -1 ? row.RowData[binningTable.Ft2GbIdx] : "";
                    if (binningRows.Any(x => x.RowData[binningTable.EqnIdx] == eq))
                    {
                        BinningRow allowModeData = binningRows.First(x => x.RowData[binningTable.EqnIdx] == eq);
                        if (!allowModeData.RowData[binningTable.CIdx].EqualsIgnoreCase(c))
                        {
                            string errorMessage = $"C value {c} is different from {allowModeData.RowData[binningTable.CIdx]} in {allowEqualMode} when setting allow equal mode!!!";
                            binningTable.AddError(BinCutErrorType.E_C_01, binningTable.SheetName, row.RowNum, binningTable.CIdx + 1, $"C value {c} is different from {allowModeData.RowData[binningTable.CIdx]} in {allowEqualMode} when setting allow equal mode!!!", [c, allowModeData.RowData[binningTable.CIdx], allowEqualMode]);
                        }

                        if (!allowModeData.RowData[binningTable.MIdx].EqualsIgnoreCase(m))
                        {
                            string errorMessage = $"M value {m} is different from {allowModeData.RowData[binningTable.MIdx]} in {allowEqualMode} when setting allow equal mode!!!";
                            binningTable.AddError(BinCutErrorType.E_M_01, binningTable.SheetName, row.RowNum, binningTable.MIdx + 1, $"M value {m} is different from {allowModeData.RowData[binningTable.MIdx]} in {allowEqualMode} when setting allow equal mode!!!", [m, allowModeData.RowData[binningTable.MIdx], allowEqualMode]);
                        }

                        if (binningTable.CpGbIdx != -1 && !allowModeData.RowData[binningTable.CpGbIdx].EqualsIgnoreCase(cpgb))
                        {
                            string errorMessage = $"CPGB value {cpgb} is different from {allowModeData.RowData[binningTable.CpGbIdx]} in {allowEqualMode} when setting allow equal mode!!!";
                            binningTable.AddError(BinCutErrorType.E_Cpgb_01, binningTable.SheetName, row.RowNum, binningTable.CpGbIdx + 1, $"CPGB value {cpgb} is different from {allowModeData.RowData[binningTable.CpGbIdx]} in {allowEqualMode} when setting allow equal mode!!!", [cpgb, allowModeData.RowData[binningTable.CpGbIdx], allowEqualMode]);
                        }

                        if (binningTable.Cp2GbIdx != -1 && !allowModeData.RowData[binningTable.Cp2GbIdx].EqualsIgnoreCase(cp2Gb))
                        {
                            string errorMessage = $"CP2GB value {cp2Gb} is different from {allowModeData.RowData[binningTable.Cp2GbIdx]} in {allowEqualMode} when setting allow equal mode!!!";
                            binningTable.AddError(BinCutErrorType.E_Cp2Gb_01, binningTable.SheetName, row.RowNum, binningTable.Cp2GbIdx + 1, $"CP2GB value {cp2Gb} is different from {allowModeData.RowData[binningTable.Cp2GbIdx]} in {allowEqualMode} when setting allow equal mode!!!", [cp2Gb, allowModeData.RowData[binningTable.Cp2GbIdx], allowEqualMode]);
                        }

                        if (binningTable.Ft1GbIdx != -1 && !allowModeData.RowData[binningTable.Ft1GbIdx].EqualsIgnoreCase(ft1Gb))
                        {
                            string errorMessage = $"FT1GB value {ft1Gb} is different from {allowModeData.RowData[binningTable.Ft1GbIdx]} in {allowEqualMode} when setting allow equal mode!!!";
                            binningTable.AddError(BinCutErrorType.E_FtRoomGb_01, binningTable.SheetName, row.RowNum, binningTable.Ft1GbIdx + 1, $"FT1GB value {ft1Gb} is different from {allowModeData.RowData[binningTable.Ft1GbIdx]} in {allowEqualMode} when setting allow equal mode!!!", [ft1Gb, allowModeData.RowData[binningTable.Ft1GbIdx], allowEqualMode]);
                        }

                        if (binningTable.Ft2GbIdx != -1 && !allowModeData.RowData[binningTable.Ft2GbIdx].EqualsIgnoreCase(ft2Gb))
                        {
                            string errorMessage = $"FT2GB value {ft2Gb} is different from {allowModeData.RowData[binningTable.Ft2GbIdx]} in {allowEqualMode} when setting allow equal mode!!!";
                            binningTable.AddError(BinCutErrorType.E_FtHotGb_01, binningTable.SheetName, row.RowNum, binningTable.Ft2GbIdx + 1, $"FT2GB value {ft2Gb} is different from {allowModeData.RowData[binningTable.Ft2GbIdx]} in {allowEqualMode} when setting allow equal mode!!!", [ft2Gb, allowModeData.RowData[binningTable.Ft2GbIdx], allowEqualMode]);
                        }
                    }
                    else
                    {
                        //can't find same equation from allow equal mode
                        string errorMessage = $"Equation {eq} not be found in {allowEqualMode} when setting allow equal mode!!!";
                        binningTable.AddError(BinCutErrorType.E_Equation_02, binningTable.SheetName, row.RowNum, 0, $"Equation {eq} not be found in {allowEqualMode} when setting allow equal mode!!!", [eq, allowEqualMode]);
                    }
                }
            }
        }

        private void CheckAllowEqual()
        {
            //Get all AllowEqual and comment mode in all sheets
            var hasAllowEqualMode = new List<string>();
            var hasCommentMode = new List<string>();
            var allowEqualList = new Dictionary<string, string>();
            foreach (BinningTable table in this)
            {
                if (table.AllowEqualIdx == -1 || table.CommentIdx == -1)
                {
                    continue;
                }

                hasAllowEqualMode.AddRange([.. table.Rows.Where(x => !string.IsNullOrEmpty(x.RowData[table.AllowEqualIdx])).Select(y => y.RowData[table.ModeIdx])]);
                hasCommentMode.AddRange([.. table.Rows.Where(x => x.RowData[table.CommentIdx].StartsWithIgnoreCase("Max PV")).Select(y => y.RowData[table.ModeIdx])]);
                foreach (BinningRow row in table.Rows)
                {
                    string mode = row.RowData[table.ModeIdx];
                    string allequal = row.RowData[table.AllowEqualIdx];
                    if (!string.IsNullOrEmpty(allequal))
                    {
                        allowEqualList.TryAdd(allequal, mode);
                    }
                }
            }
            List<List<string>> groups = GetGroups(allowEqualList);
            hasAllowEqualMode = [.. hasAllowEqualMode.Distinct()];
            hasCommentMode = [.. hasCommentMode.Distinct()];

            foreach (BinningTable table in this)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    string mode = table.Rows[i].RowData[table.ModeIdx];
                    string allequal = table.Rows[i].RowData[table.AllowEqualIdx];
                    string comment = table.Rows[i].RowData[table.CommentIdx];
                    int rowIdx = table.Rows[i].RowNum;

                    if (hasAllowEqualMode.Any(x => x.EqualsIgnoreCase(mode)))
                    {
                        if (string.IsNullOrEmpty(allequal))
                        {
                            //string errorMessage = "The Allow Equal should not be empty !!!";
                            table.AddError(BinCutErrorType.E_AllowEqual_02, table.SheetName, rowIdx, table.AllowEqualIdx + 1, "The Allow Equal should not be empty !!!");
                        }
                    }

                    if (hasCommentMode.Any(x => x.EqualsIgnoreCase(mode)))
                    {
                        if (string.IsNullOrEmpty(comment))
                        {
                            //string errorMessage = "The Comment should not be empty !!!";
                            table.AddError(BinCutErrorType.E_AllowEqual_03, table.SheetName, rowIdx, table.CommentIdx + 1, "The Comment should not be empty !!!");
                        }
                        else
                        {
                            foreach (List<string> group in groups)
                            {
                                if (group.Any(x => x.EqualsIgnoreCase(mode)))
                                {
                                    bool flag = false;
                                    foreach (string itme in group)
                                    {
                                        if (comment.IndexOf(itme, StringComparison.OrdinalIgnoreCase) < 0)
                                        {
                                            flag = true;
                                            break;
                                        }
                                    }
                                    if (flag)
                                    {
                                        string errorMessage = $"The Comment should contain mode {string.Join(",", group)} !!!";
                                        table.AddError(BinCutErrorType.E_AllowEqual_04, table.SheetName, rowIdx, table.CommentIdx + 1, $"The Comment should contain mode {string.Join(",", group)} !!!", [string.Join(",", group)]);
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        private static List<List<string>> GetGroups(Dictionary<string, string> allowEqualList)
        {
            var groups = new List<List<string>>();
            if (allowEqualList.Count == 0)
            {
                return groups;
            }

            foreach (KeyValuePair<string, string> row in allowEqualList)
            {
                string serachMode = row.Key;
                while (true)
                {
                    bool find = allowEqualList.Any(x => x.Key.EqualsIgnoreCase(serachMode));
                    if (!find)
                    {
                        break;
                    }

                    if (!groups.SelectMany(x => x).Any(x => x == serachMode))
                    {
                        groups.Add([serachMode]);
                    }

                    KeyValuePair<string, string> findRow = allowEqualList.First(x => x.Key.EqualsIgnoreCase(serachMode));
                    if (groups.SelectMany(x => x).Any(x => x == findRow.Key) && groups.SelectMany(x => x).Any(x => x == findRow.Value))
                    {
                        break;
                    }

                    foreach (List<string> group in groups)
                    {
                        if (group.Any(x => x.EqualsIgnoreCase(findRow.Key)))
                        {
                            if (!group.Any(x => x.EqualsIgnoreCase(findRow.Value)))
                            {
                                group.Add(findRow.Value);
                            }
                            serachMode = findRow.Value;
                            break;
                        }
                    }
                }
            }
            return groups;
        }

        public void CheckfuseBitDef(Dictionary<string, string> idsList, string baseVoltage, Dictionary<string, List<string>> vddList)
        {
            _ = double.TryParse(baseVoltage, out double basevoltageValue);
            var pinList = new List<string>();
            for (int idx = 0; idx < Count; idx++)
            {
                double stepSizeFromNotes = this[idx].StepSize;
                if (this[idx].IdsMaxIdx != -1)
                {
                    BinningTable vddBinTable = this[idx];
                    for (int i = 0; i < vddBinTable.Rows.Count; i++)
                    {
                        string domain = vddBinTable.Rows[i].RowData[vddBinTable.DomainIdx];
                        string mode = vddBinTable.Rows[i].RowData[vddBinTable.ModeIdx];
                        string binned = vddBinTable.Rows[i].RowData[vddBinTable.BinnedIdx];
                        int rowIdx = vddBinTable.Rows[i].RowNum;
                        #region ids
                        string isdBit = "";
                        string idsName = "";
                        bool flag = false;
                        bool flagContains = false;
                        foreach (KeyValuePair<string, string> item in idsList)
                        {
                            if (item.Key.EqualsIgnoreCase("IDS_VDD_" + domain))
                            {
                                idsName = item.Key;
                                isdBit = item.Value;
                                flag = true;
                                break;
                            }

                            if (item.Key.Contains(domain, StringComparison.OrdinalIgnoreCase))
                            {
                                idsName = item.Key;
                                isdBit = item.Value;
                                flagContains = true;
                                break;
                            }
                        }
                        if (!flag)
                        {
                            if (flagContains)
                            {
                                string errorMessage = $"The IDS_VDD_{domain} can not be found in efuseBitDef, and only found relative name {idsName} !!!";
                                vddBinTable.AddError(BinCutErrorType.E_efuseBitDef_01, vddBinTable.SheetName, rowIdx, vddBinTable.DomainIdx + 1, $"The IDS_VDD_{domain} can not be found in efuseBitDef, and only found relative name {idsName} !!!", [domain, idsName]);
                            }
                            else
                            {
                                string errorMessage = $"The IDS_VDD_{domain} can not be found in efuseBitDef !!!";
                                vddBinTable.AddError(BinCutErrorType.E_efuseBitDef_02, vddBinTable.SheetName, rowIdx, vddBinTable.DomainIdx + 1, $"The IDS_VDD_{domain} can not be found in efuseBitDef !!!", [domain]);
                            }
                        }

                        string idsMax = vddBinTable.Rows[i].RowData[vddBinTable.IdsMaxIdx];
                        _ = double.TryParse(idsMax, out double idsMaxValue);
                        _ = int.TryParse(isdBit, out int value);
                        if (idsMaxValue > Math.Pow(2, value) - 1)
                        {
                            string errorMessage = $"The IdsMax {idsMaxValue} is larger than the limit in efuseBitDef {idsName} - ( 2 ^ {isdBit} - 1 )  !!!";
                            vddBinTable.AddError(BinCutErrorType.E_IdsMax_01, vddBinTable.SheetName, rowIdx, vddBinTable.IdsMaxIdx + 1, $"The IdsMax {idsMaxValue} is larger than the limit in efuseBitDef {idsName} - ( 2 ^ {isdBit} - 1 )  !!!", [idsMaxValue.ToString(), idsName, isdBit]);
                        }

                        #endregion

                        #region product
                        bool flagVdd = false;
                        bool flagContainsVdd = false;
                        KeyValuePair<string, List<string>>? row = null;
                        if (vddList.Any(x => x.Key.EqualsIgnoreCase("VDD_" + domain + "_" + mode)))
                        {
                            flagVdd = true;
                            row = vddList.First(x => x.Key.EqualsIgnoreCase("VDD_" + domain + "_" + mode));
                        }
                        else if (vddList.Any(y => y.Key.Contains(mode, StringComparison.OrdinalIgnoreCase)))
                        {
                            flagContainsVdd = true;
                            row = vddList.First(y => y.Key.Contains(mode, StringComparison.OrdinalIgnoreCase));
                            string errorMessage = $"The VDD_{domain}_{mode} can not be found in efuseBitDef, and only found relative name {row.Value.Key} !!!";
                            vddBinTable.AddError(BinCutErrorType.W_efuseBitDef_01, vddBinTable.SheetName, rowIdx, vddBinTable.ModeIdx + 1, $"The VDD_{domain}_{mode} can not be found in efuseBitDef, and only found relative name {row.Value.Key} !!!", [domain, mode, row.Value.Key]);
                        }

                        if (!flagVdd && !flagContainsVdd && binned.EqualsIgnoreCase("True"))
                        {
                            string errorMessage = $"The VDD_{domain}_{mode} can not be found in efuseBitDef !!!";
                            vddBinTable.AddError(BinCutErrorType.E_efuseBitDef_02, vddBinTable.SheetName, rowIdx, vddBinTable.ModeIdx + 1, $"The VDD_{domain}_{mode} can not be found in efuseBitDef !!!", [domain, mode]);
                        }

                        if (row != null)
                        {
                            string productBit = row.Value.Value[3];
                            _ = int.TryParse(productBit, out int productBitvalue);
                            string resolution = row.Value.Value[11];
                            _ = double.TryParse(resolution, out double resolutionValue);
                            if (stepSizeFromNotes.CompareTo(resolutionValue) != 0)
                            {
                                if (!pinList.Exists(x => x.EqualsIgnoreCase(row.Value.Value[0])))
                                {
                                    string errorMessage = $"The efuse {row.Value.Value[0]} resolution {resolutionValue} is different with StepSize {stepSizeFromNotes} from the Notes sheet !!!";
                                    vddBinTable.AddError(BinCutErrorType.E_Resolution_01, vddBinTable.SheetName, 2, 3, $"The efuse {row.Value.Value[0]} resolution {resolutionValue} is different with StepSize {stepSizeFromNotes} from the Notes sheet !!!", [row.Value.Value[0], resolutionValue.ToString(), stepSizeFromNotes.ToString()]);
                                    pinList.Add(row.Value.Value[0]);
                                }
                            }

                            string cpVmax = vddBinTable.Rows[i].RowData[vddBinTable.CpVMaxIdx];
                            _ = double.TryParse(cpVmax, out double cpVmaxValue);

                            string cpGb = vddBinTable.Rows[i].RowData[vddBinTable.CpGbIdx];
                            _ = double.TryParse(cpGb, out double cpGbValue);

                            if (cpVmaxValue + cpGbValue - basevoltageValue > (Math.Pow(2, productBitvalue) - 1) * resolutionValue)
                            {
                                string errorMessage = $"The cpVMax {cpVmaxValue} + {cpGb} - {baseVoltage} is larger than the limit in efuseBitDef {mode} - ( 2 ^ {productBit} - 1 ) * {resolutionValue} !!!";
                                vddBinTable.AddError(BinCutErrorType.E_CpMax_01, vddBinTable.SheetName, rowIdx, vddBinTable.CpVMaxIdx + 1, $"The cpVMax {cpVmaxValue} + {cpGb} - {baseVoltage} is larger than the limit in efuseBitDef {mode} - ( 2 ^ {productBit} - 1 ) * {resolutionValue} !!!", [cpVmaxValue.ToString(), cpGb, baseVoltage, mode, productBit, resolutionValue.ToString()]);
                            }
                        }
                        #endregion
                    }
                }
            }
        }

        public void CheckDomain(IEnumerable<string> powerPins)
        {
            var domains = new List<string>();
            foreach (string powerPin in powerPins)
            {
                string pinName = powerPin.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries).First();
                List<string> arr = [.. pinName.Split('_')];
                arr.RemoveAt(0);
                domains.Add(string.Join("_", arr));
            }

            var dic = new Dictionary<string, string>();
            for (int idx = 0; idx < Count; idx++)
            {
                BinningTable vddBinTable = this[idx];
                for (int i = 0; i < vddBinTable.Rows.Count; i++)
                {
                    if (this[idx].DomainIdx != -1 && this[idx].ModeIdx != -1)
                    {
                        string domain = vddBinTable.Rows[i].RowData[vddBinTable.DomainIdx];
                        string mode = vddBinTable.Rows[i].RowData[vddBinTable.ModeIdx];
                        dic.TryAdd(mode, domain);
                    }
                }
            }

            for (int idx = 0; idx < Count; idx++)
            {
                BinningTable vddBinTable = this[idx];
                for (int i = 0; i < vddBinTable.Rows.Count; i++)
                {
                    if (this[idx].DomainIdx != -1)
                    {
                        string domain = vddBinTable.Rows[i].RowData[vddBinTable.DomainIdx];
                        string mode = vddBinTable.Rows[i].RowData[vddBinTable.ModeIdx];
                        if (this[idx].Rows[i].RowData[this[idx].BinnedIdx].EqualsIgnoreCase("True") ||
                            this[idx].Rows[i].RowData[this[idx].BinnedIdx].EqualsIgnoreCase("False") ||
                            this[idx].Rows[i].RowData[this[idx].BinnedIdx].EqualsIgnoreCase("ATE"))
                        {
                            int rowIdx = vddBinTable.Rows[i].RowNum;
                            if (!domains.Exists(x => x.EqualsIgnoreCase(domain)))
                            {
                                string errorMessage = $"The domain {domain} is not existed !!!";
                                vddBinTable.AddError(BinCutErrorType.E_Domain_01, vddBinTable.SheetName, rowIdx, vddBinTable.DomainIdx + 1, $"The domain {domain} is not existed !!!", [domain]);
                            }

                            string domainInSameMode = "";
                            if (dic.TryGetValue(mode, out string? value) && !string.IsNullOrEmpty(value))
                            {
                                domainInSameMode = value;
                            }

                            if (!string.IsNullOrEmpty(domainInSameMode) && !domain.EqualsIgnoreCase(domainInSameMode))
                            {
                                string errorMessage = $"The domain {domain} of {mode} is correct (Should be {domainInSameMode}) !!!";
                                vddBinTable.AddError(BinCutErrorType.E_Domain_02, vddBinTable.SheetName, rowIdx, vddBinTable.DomainIdx + 1, $"The domain {domain} of {mode} is correct (Should be {domainInSameMode}) !!!", [domain, mode, domainInSameMode]);
                            }
                        }
                    }
                }
            }
        }
        #endregion

        public void InsertMissingHeader(ExcelWorksheet excelWorksheet)
        {
            var titleList = this.SelectMany(x => x.TitleList).Where(x => !string.IsNullOrEmpty(x.Item1)).Distinct().ToList();
            BinningTable? sheet = Find(x => x.SheetName.EqualsIgnoreCase(excelWorksheet.Name));
            if (sheet != null)
            {
                var titles = titleList.Select(x => x.Item1).Except(sheet.TitleList.Select(x => x.Item1), StringExtensions.IgnoreCase).ToList();
                if (titles.Count != 0)
                {
                    Tuple<string, int>? comment = sheet.TitleList.Find(x => x.Item1.EqualsIgnoreCase("Comment"));
                    int insertColumn = excelWorksheet.Dimension.End.Column + 1;
                    if (comment != null)
                    {
                        insertColumn = comment.Item2 + 1;
                    }

                    foreach (string title in titles)
                    {
                        excelWorksheet.InsertColumn(title, sheet.StartRowIdx + 1, insertColumn);
                    }
                }
            }
        }

        private void CheckIDvalue()
        {
            //1.not duplicate in the same domain
            //2.id value need to greater than id value of the next euqation or next performance mode
            foreach (BinningTable table in this)
            {
                if (table.IdIdx == -1)
                {
                    // ID column not exist
                    string errorMessage = $"The ID column does not exist in sheet {table.SheetName}, please add it!!!";
                    table.AddError(BinCutErrorType.E_ID_01, table.SheetName, 0, 0, $"The ID column does not exist in sheet {table.SheetName}, please add it!!!", [table.SheetName]);
                    continue;
                }
                if (table.DomainIdx == -1 || table.ModeIdx == -1)
                {
                    continue;
                }

                var domainGroups = table.Rows.GroupBy(x => x.RowData[table.DomainIdx] + "_" + x.RowData[table.ModeIdx][..2]).ToList();
                foreach (IGrouping<string, BinningRow> domainGroup in domainGroups)
                {
                    var pmodeGroups = domainGroup.GroupBy(x => x.RowData[table.ModeIdx]).ToList();
                    string tmpLastEq = "";
                    double tmpLastId = 0;
                    string tmpPmode = "";
                    foreach (IGrouping<string, BinningRow> pmodeGroup in pmodeGroups)
                    {
                        #region check id value of each equation in same performance mode
                        string tmpEqOri = pmodeGroup.ElementAt(0).RowData[table.EqnIdx];
                        int tmpEq = Convert.ToInt32(tmpEqOri.Replace("E", ""));

                        if (!_regex.IsMatch(pmodeGroup.ElementAt(0).RowData[table.IdIdx]))
                        {
                            //string errorMessage = "The ID value incorrect";
                            table.AddError(BinCutErrorType.E_ID_02, table.SheetName, pmodeGroup.ElementAt(0).RowNum, table.IdIdx + 1, "The ID value invalid");
                            continue;
                        }
                        double tmpId = Convert.ToDouble(pmodeGroup.ElementAt(0).RowData[table.IdIdx]);
                        for (int i = 1; i < pmodeGroup.Count(); i++)
                        {
                            string eqOri = pmodeGroup.ElementAt(i).RowData[table.EqnIdx];
                            int eq = Convert.ToInt32(eqOri.Replace("E", ""));
                            if (!_regex.IsMatch(pmodeGroup.ElementAt(i).RowData[table.IdIdx]))
                            {
                                //string errorMessage = "The ID value incorrect";
                                table.AddError(BinCutErrorType.E_ID_02, table.SheetName, pmodeGroup.ElementAt(i).RowNum, table.IdIdx + 1, "The ID value invalid");
                                continue;
                            }
                            double id = Convert.ToDouble(pmodeGroup.ElementAt(i).RowData[table.IdIdx]);
                            if (!(tmpEq < eq && tmpId < id) || !(tmpEq > eq && tmpId > id))
                            {
                                if (tmpEq < eq && tmpId > id)
                                {
                                    string errorMessage = $"The larger equation {eqOri} ({id}) needs to set a larger ID value than {tmpEqOri} ({tmpId})";
                                    table.AddError(BinCutErrorType.E_ID_04, table.SheetName, pmodeGroup.ElementAt(i).RowNum, table.IdIdx + 1, $"The larger equation {eqOri} ({id}) needs to set a larger ID value than {tmpEqOri} ({tmpId})", [eqOri, id.ToString(), tmpEqOri, tmpId.ToString()]);
                                }

                                if (tmpEq > eq && tmpId < id)
                                {
                                    string errorMessage = $"The smaller equation {eqOri} ({id}) needs to set a smaller ID value than {tmpEqOri} ({tmpId})";
                                    table.AddError(BinCutErrorType.E_ID_05, table.SheetName, pmodeGroup.ElementAt(i).RowNum, table.IdIdx + 1, $"The smaller equation {eqOri} ({id}) needs to set a smaller ID value than {tmpEqOri} ({tmpId})", [eqOri, id.ToString(), tmpEqOri, tmpId.ToString()]);
                                }
                            }
                            tmpEqOri = eqOri;
                            tmpEq = eq;
                            tmpId = id;
                        }
                        #endregion

                        #region check id value of the last equation in different performance mode

                        if (string.IsNullOrEmpty(tmpPmode) && string.IsNullOrEmpty(tmpLastEq))
                        {
                            tmpLastEq = pmodeGroup.Last().RowData[table.EqnIdx];
                            tmpLastId = Convert.ToDouble(pmodeGroup.Last().RowData[table.IdIdx]);
                            tmpPmode = pmodeGroup.Last().RowData[table.ModeIdx];
                        }
                        else
                        {
                            string lastEq = pmodeGroup.Last().RowData[table.EqnIdx];
                            double lastId = Convert.ToDouble(pmodeGroup.Last().RowData[table.IdIdx]);
                            string pmode = pmodeGroup.Last().RowData[table.ModeIdx];
                            if (tmpLastId >= lastId)
                            {
                                string errorMessage = $"The {lastEq}(ID:{lastId}) of the performance mode {pmode} needs to be set to be greater than the {tmpLastEq}(ID:{tmpLastId}) of the {tmpPmode}";
                                table.AddError(BinCutErrorType.E_ID_06, table.SheetName, pmodeGroup.Last().RowNum, table.IdIdx + 1, $"The {lastEq}(ID:{lastId}) of the performance mode {pmode} needs to be set to be greater than the {tmpLastEq}(ID:{tmpLastId}) of the {tmpPmode}", [lastEq, lastId.ToString(), pmode, tmpLastEq, tmpLastId.ToString(), tmpPmode]);
                            }
                            tmpLastEq = lastEq;
                            tmpLastId = lastId;
                            tmpPmode = pmode;
                        }
                        #endregion

                    }

                    #region check if ID value is duplicated in the same domain
                    var idGroups = domainGroup.GroupBy(x => x.RowData[table.IdIdx]).ToList();
                    foreach (IGrouping<string, BinningRow> idGroup in idGroups)
                    {
                        if (idGroup.Count() > 1)
                        {
                            // ID value duplicate
                            string errorMessage = $"The ID value is duplicated in the {idGroup.First().RowData[table.DomainIdx]} domain, please change!!!";
                            table.AddError(BinCutErrorType.E_ID_07, table.SheetName, idGroup.First().RowNum, table.IdIdx + 1, $"The ID value is duplicated in the {idGroup.First().RowData[table.DomainIdx]} domain, please change!!!", [idGroup.First().RowData[table.DomainIdx]]);
                        }
                    }
                    #endregion
                }
            }
        }

        public List<string> GetBinningTitleList()
        {
            var binningTitleList = new List<string>();

            foreach (BinningTable binning in this)
            {
                binningTitleList.AddRange(binning.TitleList.Select(x => x.Item1));
            }

            return [.. binningTitleList.Distinct()];
        }

        public Dictionary<string, List<string>> GetDomainDic()
        {
            var domainDic = new Dictionary<string, List<string>>();
            foreach (BinningRow row in this[0].Rows)
            {
                string domain = row.RowData[this[0].DomainIdx];
                string mode = row.RowData[this[0].ModeIdx];
                if (!domainDic.ContainsKey(domain))
                {
                    domainDic.Add(domain, [mode]);
                }
                else
                {
                    if (!domainDic[domain].Exists(x => x == mode))
                    {
                        domainDic[domain].Add(mode);
                    }
                }
            }
            return domainDic;
        }
    }
}
