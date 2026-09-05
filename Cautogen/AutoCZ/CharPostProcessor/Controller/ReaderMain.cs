using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.InputReader;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.common.ReaderWriter.Reader.InputReader;

using LogLib.Utility;

using TestPlanLib.PatternListCsvFile;

using HardIpReference = Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.HardIpReference;
using PatInfoReader = Cautogen.AutoCZ.CharPostProcessor.InputReader.PatInfoReader;

namespace Cautogen.AutoCZ.CharPostProcessor.Controller
{
    public class ReaderMain
    {
        public static void Run(InputParam param)
        {
            try
            {
                LogHelper.Info("Reading PatInfo ...");
                if (!string.IsNullOrEmpty(param.PatInfoFile))
                {
                    LocalSpecs.PatInfoList = PatInfoReader.Read(param.PatInfoFile).ToDictionary(a => a.Payload.ToLower(), a => a);
                    if (param.HardIpInfoAllDict == null)
                    {
                        param.HardIpInfoAllDict = ConvertInfoSubRoutine(LocalSpecs.PatInfoList);
                    }
                }
                if (File.Exists(param.PatListFile))
                {
                    new PatternListInputReader(param.PatListFile).Read();
                    LocalSpecs.PatternDatas = PatternListInputReader.PatternList;
                }
                else
                {
                    LocalSpecs.PatternDatas = CharPlanReader.ReadPatternDashBoard(param.CharFile);
                }
                //LocalSpecs.EmaMappingItems = new ().ReadMappingTable(param.CharFile);

                foreach (KeyValuePair<string, PatternData> row in LocalSpecs.PatternDatas)
                {
                    string timeset = row.Value.TimesetVersion;
                    if (param.TimeSetVersionDic.TryGetValue(timeset, out TimeSetItem value))
                    {
                        row.Value.TimesetVersion = timeset + "_" + value.Version;
                    }
                }
                LogHelper.Info("Reading Char Plan ...");
                LocalSpecs.CharPlanSheets = CharPlanReader.Read(param.CharFile);

            }
            catch (Exception e)
            {
                GeneralFunc.WriteMessage("Reading input files failed " + e.Message);
            }
        }

        private static Dictionary<string, SubrPatInfo> ConvertInfoSubRoutine(Dictionary<string, HardIpReference> infos)
        {
            var result = new Dictionary<string, SubrPatInfo>();
            foreach (KeyValuePair<string, HardIpReference> info in infos)
            {
                var subr = new SubrPatInfo { Subroutine = new List<string> { info.Value.Subr }, VmVector = info.Value.Vm };
                result.Add(info.Key, subr);
            }
            return result;
        }
    }
}
