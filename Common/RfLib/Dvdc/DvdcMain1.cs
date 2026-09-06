using System;
using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace RfLib.Dvdc
{
    public class DvdcMain1 : HardIpMain
    {
        public const string CharSheetName = "Char_Wireless";

        public override void WorkFlow(HardIpParaData hardIpParaData)
        {
            try
            {
                (List<TimeSetBasicSheet> timeSetSheets, PatSetSheet patSetSheet, WaveDefinitionSheet waveDefinitionSheet, MixedSignalSheet mixedSignalSheet, List<InstanceSheet> instSheets, List<SubFlowSheet> flowSheets, BinTableSheet binTableSheet, CharSheet charSheet, SubFlowSheet initFlow) = DvdcSheetsGenerator.Generate(HardIpInputData);

                #region Add result sheets to result workbook
                //Add Init Flag Flow
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, initFlow);

                //Add SubFlow
                foreach (SubFlowSheet flowSheet in flowSheets)
                {
                    //Delete jobs if All jobs are enable and replace job name in job column by actual job in config
                    flowSheet.FilterFlowJobs(LocalSpecs.AllJobsHardIp);
                    if (flowSheet.Rows.Count > 0)
                    {
                        TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, flowSheet);
                    }
                }

                //Add Instance sheet
                foreach (InstanceSheet instSheet in instSheets)
                {
                    if (instSheet.Rows.Count > 0)
                    {
                        TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirHardIp, instSheet);
                    }
                }

                //Add Bin_Table_HardIP
                TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirBinTable, binTableSheet);

                //Add WaveDef
                if (waveDefinitionSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddWaveDefSheet(FolderStructure.DirWaveDef, waveDefinitionSheet);
                }

                //Add Mixed Signal
                if (mixedSignalSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddMixedSignalSheet(FolderStructure.DirMixedSignal, mixedSignalSheet);
                }

                //Add char
                if (charSheet.Rows.Count > 0)
                {
                    TestProgram.IgxlWorkBk.AddCharSheet(FolderStructure.DirDevChar, charSheet);
                }

                //Add timeSet
                foreach (TimeSetBasicSheet timeSetSheet in timeSetSheets)
                {
                    TestProgram.IgxlWorkBk.AddTimeSetSheet(FolderStructure.DirTimings, timeSetSheet);
                }

                //Add PatSet
                if (patSetSheet.Rows.Count != 0)
                {
                    TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirPatSetsAll, patSetSheet);
                }
                #endregion

            }
            catch (Exception e)
            {
                string outString = string.Format("LCD AutoGen got a fatal mistake! \n" + e.StackTrace);
                Response.Report(outString, EnumMessageLevel.Error, 10);
            }
        }
    }
}
