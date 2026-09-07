using System;
using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader;

using CommonLib.Enums;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    public class InputDefReader : ExcelReader
    {
        /* properites */
        public static Dictionary<string, InputDefRow> BlockMapDict = new Dictionary<string, InputDefRow>(); // char plan block name => DC Category

        /* constructor */

        public InputDefReader(string filePath, Action callbackFunc = null)
            : base(filePath, callbackFunc)
        {
            BlockMapDict = new Dictionary<string, InputDefRow>();
        }

        /* methods */
        protected override void _Read(ExcelWorksheet sh)
        {
            // read block mapping
            if (sh.Name != "Block Mapping")
            {
                return;
            }

            MessageWriter.WriteMessage("Process Char_Input_Def file...", EnumMessageLevel.Info);

            // parse header
            var headerColDict = new Dictionary<string, int>();
            for (int i = 1; i <= sh.Dimension.End.Column; i++)
            {
                if (sh.Cells[1, i].Value == null)
                {
                    continue;
                }

                string header = sh.Cells[1, i].Value.ToString().Replace(" ", "");
                headerColDict.Add(header, i);
            }

            int blockCol = _GetHeaderCol(headerColDict, "CharPlanBlockName");
            int groupCol = _GetHeaderCol(headerColDict, "CharPlanGroup");
            int dcCatCol = _GetHeaderCol(headerColDict, "DCCategory");
            int acCatCol = _GetHeaderCol(headerColDict, "ACCategory");
            int levelCol = _GetHeaderCol(headerColDict, "LevelSheet");
            int timeSetCol = _GetHeaderCol(headerColDict, "TimingSheet");
            int pModeCol = _GetHeaderCol(headerColDict, "DCSpecwithPerformanceMode");
            int powerRunIndex = _GetHeaderCol(headerColDict, "PowerRunRatio", "PowerRunScenario");

            // parse content
            for (int i = 2; i <= sh.Dimension.End.Row; i++)
            {
                if (sh.Cells[i, blockCol].Value == null)
                {
                    continue;
                }

                var blockMap = new InputDefRow
                {
                    BlockName = sh.Cells[i, blockCol].Value.ToString().ToUpper(),
                    Group = sh.Cells[i, groupCol].Value == null ? "" : sh.Cells[i, groupCol].Value.ToString().ToUpper(),
                    DcCategory = sh.Cells[i, dcCatCol].Value == null ? "" : sh.Cells[i, dcCatCol].Value.ToString().ToUpper(),
                    AcCategory = sh.Cells[i, acCatCol].Value == null ? "" : sh.Cells[i, acCatCol].Value.ToString(),
                    LevelSheet = sh.Cells[i, levelCol].Value == null ? "" : sh.Cells[i, levelCol].Value.ToString(),
                    TimingSheet = sh.Cells[i, timeSetCol].Value == null ? "" : sh.Cells[i, timeSetCol].Value.ToString(),
                    UsePMode = sh.Cells[i, pModeCol].Value != null && sh.Cells[i, pModeCol].Value.ToString().Trim() == "Y",
                    PowerRunScenario = sh.Cells[i, powerRunIndex].Value == null ? "" : sh.Cells[i, powerRunIndex].Value.ToString()
                };

                // update to UtilityData
                if (string.IsNullOrEmpty(blockMap.Group))
                {
                    BlockMapDict[blockMap.BlockName] = blockMap;
                }
                else
                {
                    BlockMapDict[blockMap.BlockName + ":" + blockMap.Group] = blockMap;
                }
            }
        }

        protected override void _ReadPreProcess(ExcelWorksheet sh)
        {
            //base._Preprocess(sh);
        }

        private static int _GetHeaderCol(IDictionary<string, int> headerOrder, string header, string headerAlies = null)
        {
            if (headerOrder.TryGetValue(header, out int col))
            {
                return col;
            }

            if (headerAlies != null && headerOrder.TryGetValue(headerAlies, out int headerCol))
            {
                return headerCol;
            }

            throw new Exception("Missing header " + header);
        }
    }
}
