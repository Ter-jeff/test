using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using RfLib.InstrumentSetup;
using RfLib.InstrumentSetup.InstrumentTypeData;

using DataTable = System.Data.DataTable;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public partial class GenInstrumentSetupTable
    {
        [GeneratedRegex(@"(.*_SI_(?<value>.*))|(.*_DM_(?<value>.*))", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"^\[(?<setup>.*)\]$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        public InstrumentSetupWriter Writer { get; set; } = new InstrumentSetupWriter();
        public DataTable DtInstrumentSetup { get; set; } = new DataTable();
        public List<InstrumentSetupRow> InstrumentSetupRowList { get; set; } = [];
        private int _rowNumber = 3;

        public void Init(List<ChannelMapSheet> channelMapSheets)
        {
            string basicConfig = string.Format("{0}\\Settings\\Basic\\Basic_Configure_{1}.xlsx", LocalSpecs.SettingFolder, LocalSpecs.CurrentProject);
            if (!File.Exists(basicConfig))
            {
                //Copy from default
                File.Copy(Directory.GetCurrentDirectory() + "\\Settings\\Basic\\Basic_Configure_Default_RF.xlsx", basicConfig);
            }

            if (File.Exists(basicConfig))
            {
                var inputExcel = new ExcelPackage(new FileInfo(basicConfig));
                ExcelWorksheet worksheet_InstSet = inputExcel.Workbook.Worksheets["InstrumentSetup"];

                Writer = new InstrumentSetupWriter(channelMapSheets);
                Writer.ReadConfigSheet(worksheet_InstSet);
                DtInstrumentSetup = Writer.WriteDefaultHeader();
                InstrumentSetupRowList = [];

                ExcelWorksheet worksheet_ReplaceInst = inputExcel.Workbook.Worksheets["ReplaceInstrumentKey"];
                Writer.ReadReplaceConfigSheet(worksheet_ReplaceInst);
            }
        }

        public void WriteInstTableNew(HardIpInfoNew hardIpInfoNew, string patName, string scghName)
        {
            string setupName = MyRegex().Match(patName).Groups["value"].ToString();
            if (string.IsNullOrEmpty(setupName))
            {
                setupName = scghName;
            }

            int seqi = -1;
            foreach (HardIpSeqInfoNew seq in hardIpInfoNew.SeqInfo)
            {
                seqi++;

                for (int k = 0; k < seq.MeasSeq.Split('>').Length; k++)
                {
                    try
                    {
                        var rowData = new InstrumentSetupRow
                        {
                            InstrumentDataTable = DtInstrumentSetup,
                            Pattern = patName,
                            Pin = seq.MeasPins[k], ///////////////////////////////
                            Pins = string.Join(",", seq.MeasPin),
                            SeqIndex = seqi,
                            SubsetIndex = k,
                            SetupName = setupName,
                            MeasSeqType = seq.MeasSeq.Split('>')[k]
                        };

                        string rfInstrumentSetup = hardIpInfoNew.RfSetup.Split('|')[seqi].Split('>')[k];
                        string instSetupInfo =
                            MyRegex1().Match(rfInstrumentSetup)
                                .Groups["setup"]
                                .ToString();
                        //if (string.IsNullOrEmpty(instSetupInfo))
                        //    continue;
                        if (rowData.MeasSeqType.EqualsIgnoreCase("calc"))
                        {
                            continue;
                        }

                        var dicInfoPara = new Dictionary<string, string>();
                        foreach (string item in instSetupInfo.Split('$'))
                        {
                            if (string.IsNullOrEmpty(item))
                            {
                                continue;
                            }

                            string header = item.Split('=')[0].ToLower();
                            string value = item.Split('=')[1].ToLower();
                            if (header == "rlevl")
                            {
                                value = value.Replace("dbm", "");
                            }

                            dicInfoPara.Add(header.ToLower(), value);
                        }

                        //Get instrument type
                        List<InstrumentTypeSetting> typeList = Writer.CheckTypeByFreqNew(seq.MeasSeq.Split('>')[k], dicInfoPara, rowData.Pin.PinName);

                        MeasPin measPin = seq.MeasPins[k];
                        measPin.SubSequenceIndex = k;
                        List<string> instrumentTypeList = Writer.GetInstrumentType(measPin, typeList);

                        //temporary
                        if (instrumentTypeList.Exists(p => p == "UltraWave24 MultiTone Source"))
                        {
                            instrumentTypeList.Remove("UltraWave24 MultiTone Source");
                        }

                        if (instrumentTypeList.Exists(p => p == "UltraWave24 Modulated Source"))
                        {
                            instrumentTypeList.Remove("UltraWave24 Modulated Source");
                        }
                        //

                        string insType = instrumentTypeList.Count > 0 ? string.Join(",", instrumentTypeList) : "TBD";

                        if (dicInfoPara.TryGetValue("freq", out string? freqValue))
                        {
                            string[] freqList = freqValue.Split(';');
                            for (int j = 0; j < freqList.Length; j++)
                            {
                                rowData = new InstrumentSetupRow
                                {
                                    InstrumentDataTable = DtInstrumentSetup,
                                    Pattern = patName,
                                    Pin = seq.MeasPins[k + j],
                                    Pins = string.Join(",", seq.MeasPin).Split('>')[k],
                                    SeqIndex = seqi,
                                    SetupName = setupName,
                                    MeasSeqType = seq.MeasSeq.Split('>')[k],
                                    SubsetIndex = j + k
                                };

                                Writer.WriteRowData(rowData, ref _rowNumber, instrumentTypeList, dicInfoPara, insType, j + k);
                                rowData.RowIndex = _rowNumber;
                                InstrumentSetupRowList.Add(rowData);
                                _rowNumber++;
                            }
                        }
                        else
                        {
                            Writer.WriteRowData(rowData, ref _rowNumber, instrumentTypeList, dicInfoPara, insType);
                            rowData.RowIndex = _rowNumber;
                            InstrumentSetupRowList.Add(rowData);
                            _rowNumber++;
                        }
                        ;
                    }
                    catch (Exception)
                    {
                        ;
                    }
                }
            }
        }
    }
}
