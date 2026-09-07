using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;

namespace TestPlanLib.BinCut
{
    public class IdsDistribution
    {
        #region Field
        private readonly List<List<string>> _vddBinRowData;
        private readonly List<string> _modeList;
        private readonly string _outputFolder;
        private readonly int _modeid;
        private readonly int _eqnIdx;
        private readonly int _idsMaxIdx;
        private readonly string _version;
        #endregion

        public IdsDistribution(string pFolder, BinningTable binningTable)
        {
            _outputFolder = pFolder;
            _vddBinRowData = [.. binningTable.Rows.Select(x => x.RowData).Where(x => x[binningTable.BinnedIdx].EqualsIgnoreCase("true"))];
            _modeList = [];
            _modeid = binningTable.ModeIdx;
            _eqnIdx = binningTable.EqnIdx;
            _idsMaxIdx = binningTable.IdsMaxIdx;
            _version = binningTable.Version;
            _modeList = [.. _vddBinRowData.Select(x => x[_modeid]).Distinct()];
        }

        public string WorkFlow(EqnStartSheet? eqnStartSheet = null)
        {
            string[] blockList = ["TD", "MBIST", "SPI", "TMPS", "LDCBFD"];
            string destFile = Path.Combine(_outputFolder, BinCutConst.ConIdsDistribution + ".TXT");
            if (File.Exists(destFile))
            {
                File.Delete(destFile);
            }

            int modeCnt = _modeList.Count;

            using (var sw = new StreamWriter(destFile))
            {
                string line = "";
                line = string.Format("Rev:\t" + _version);
                sw.WriteLine(line);

                line = "";
                foreach (string block in blockList)
                {
                    line += string.Format(block + "\t\t");
                }

                sw.WriteLine(line);

                for (int i = 0; i < modeCnt; i++)
                {
                    line = "";
                    for (int j = 0; j < blockList.Length; j++)
                    {
                        line += string.Format(_modeList[i] + "\t\t");
                    }

                    sw.WriteLine(line);

                    line = "";
                    for (int j = 0; j < blockList.Length; j++)
                    {
                        line += "IDS Range\tStart Bin\t";
                    }

                    sw.WriteLine(line);

                    string maxEqNum = GetMaxEqNum(_modeList[i], eqnStartSheet);

                    line = "";
                    for (int j = 0; j < blockList.Length; j++)
                    {
                        line += string.Format("0\t" + maxEqNum + "\t");
                    }

                    sw.WriteLine(line);

                    var cpidsMaxList = new List<double>();
                    cpidsMaxList = [.. _vddBinRowData.Where(x => x[_modeid] == _modeList[i]).Select(y => double.Parse(y[_idsMaxIdx])).Distinct()];

                    for (int k = 0; k < cpidsMaxList.Count; k++)
                    {
                        line = "";
                        if (cpidsMaxList[k] == cpidsMaxList.Max())
                        {
                            for (int j = 0; j < blockList.Length; j++)
                            {
                                line += string.Format(cpidsMaxList[k] + "\t0\t");
                            }

                            sw.WriteLine(line);
                        }
                        else
                        {
                            for (int j = 0; j < blockList.Length; j++)
                            {
                                line += string.Format(cpidsMaxList[k] + "\t" + maxEqNum + "\t");
                            }

                            sw.WriteLine(line);
                        }
                    }

                    if (i != modeCnt - 1)
                    {
                        line = "";
                        sw.WriteLine(line);
                    }
                }
                line = "";
                for (int j = 0; j < blockList.Length; j++)
                {
                    line += "END\tEND\t";
                }

                sw.WriteLine(line);
            }

            return Path.Combine(_outputFolder, BinCutConst.ConIdsDistribution + ".txt");
        }

        public string GetMaxEqNum(string mode, EqnStartSheet? eqnStartSheet)
        {
            int maxEqNum = 0;
            int tempEqNum = 0;
            foreach (List<string> item in _vddBinRowData)
            {
                if (item[_modeid].EqualsIgnoreCase(mode))
                {
                    tempEqNum = int.Parse(item[_eqnIdx][1..]);
                }
                maxEqNum = tempEqNum > maxEqNum ? tempEqNum : maxEqNum;
            }

            if (eqnStartSheet != null)
            {
                if (eqnStartSheet.Rows.Exists(x => x.Mode == mode))
                {
                    EqnStartRow row = eqnStartSheet.Rows.Find(x => x.Mode == mode)!;
                    string eqnstart = row.Eqnstart;
                    if (!string.IsNullOrEmpty(eqnstart))
                    {
                        if (int.TryParse(eqnstart, out int eqnstartValue) &&
                            eqnstartValue > maxEqNum)
                        {
                            ErrorReportManager.AddError(BinCutErrorType.E_Pattern_02, eqnStartSheet.SheetName, row.RowNum, 0, $"The eqn start {eqnstart} can not exceed the max eqn {maxEqNum} !!!", [eqnstart, maxEqNum.ToString()]);
                        }
                        return eqnstart;
                    }
                }
            }
            return maxEqNum.ToString();
        }
    }
}
