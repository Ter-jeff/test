using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows;
using Automation.Singleton;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using ProjectConfigLib.ProjectConfig;

using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business
{
    public class LevelSheetsGenerator
    {
        #region Field
        private readonly MultiTestSettingSheetsSingleton _testSettingSheetsSingleton;
        private readonly PowerInfoSheet _powerInfo;
        private readonly PinMapSheet _pinMapSheet;
        private readonly bool _generateVSlewRate;
        private bool _hasSpiPwr;
        private MultiLevelSheets _multiLevelSheets;
        protected bool UseSaveVoltage;
        private readonly PowerLevelsGenerator _powerLevelsGenerator;
        private readonly IoLevelsGenerator _ioLevelsGenerator;
        private readonly CustomIoLevelsGenerator _customIoLevelsGenerator;
        private readonly NwireLevelsGenerator _nwireLevelsGenerator;
        private readonly RtosLevelsGenerator _rtosLevelsGenerator;
        #endregion

        #region Constructor
        public LevelSheetsGenerator(MultiTestSettingSheetsSingleton testSettingSheetsSingleton, PowerInfoSheet powerInfo, PinMapSheet pinMapSheet)
        {
            _testSettingSheetsSingleton = testSettingSheetsSingleton;
            _powerInfo = powerInfo;
            _pinMapSheet = pinMapSheet;
            _generateVSlewRate = ProjectConfigSingleton.Instance().GetValue("Basic", "Generate VSlewRate").Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            _powerLevelsGenerator = new PowerLevelsGenerator(_testSettingSheetsSingleton, _powerInfo, _pinMapSheet, _generateVSlewRate, UseSaveVoltage);
            _ioLevelsGenerator = new IoLevelsGenerator();
            _customIoLevelsGenerator = new CustomIoLevelsGenerator(_pinMapSheet);
            _nwireLevelsGenerator = new NwireLevelsGenerator();
            _rtosLevelsGenerator = new RtosLevelsGenerator();
            SetupSpi();
        }
        #endregion
        private void SetupSpi()
        {
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            _hasSpiPwr = false;
            if (pinMap.IsPinExist(CommonConst.SpiRomPwr))
            {
                _hasSpiPwr = true;
            }
            else if (pinMap.IsPinExist(CommonConst.RtosRomPwr))
            {
                _hasSpiPwr = true;
                CommonConst.SpiRomPwr = CommonConst.RtosRomPwr;
            }
        }

        public MultiLevelSheets GenerateFlow(Dictionary<string, LevelData> levelSheetData, HardIpDcSheet hardIpDcSheet = null, IoContiSheet ioContiSheet = null)
        {
            _multiLevelSheets = new MultiLevelSheets();

            CreateLevelSheets(levelSheetData).ForEach(x => _multiLevelSheets.AddLeveSheet(x));

            if (NwireSingleton.Instance().HasNwirePin)
            {
                LevelSheet nWireLevel = CreateNwireLevelSheet();
                _multiLevelSheets.AddLeveSheet(nWireLevel);
            }

            if (_hasSpiPwr)
            {
                LevelSheet spiRomLevel = CreateSpiRomLevelSheet();
                _multiLevelSheets.AddLeveSheet(spiRomLevel);
            }

            if (_multiLevelSheets.ContainsKey("Levels_HardIP"))
            {
                CreateHardIpDcLevels(_multiLevelSheets["Levels_HardIP"], hardIpDcSheet, ioContiSheet)
                    .ForEach(x => _multiLevelSheets.AddLeveSheet(x));
            }

            return _multiLevelSheets;
        }

        private List<LevelSheet> CreateLevelSheets(Dictionary<string, LevelData> levelSheetData)
        {
            List<LevelSheet> levelSheets = new List<LevelSheet>();

            foreach (KeyValuePair<string, LevelData> singleLevelSheetData in levelSheetData)
            {
                var levelSheet = new LevelSheet(singleLevelSheetData.Key);
                var tempNwireCustomLevelSheet = new LevelSheet("Levels_TempNwireCustom");
                _powerLevelsGenerator.UpdateLevelSheet(ref levelSheet, singleLevelSheetData.Value);
                _ioLevelsGenerator.UpdateLevelSheet(ref levelSheet, singleLevelSheetData.Value);
                _nwireLevelsGenerator.UpdateLevelSheet(ref tempNwireCustomLevelSheet);
                _customIoLevelsGenerator.UpdateLevelSheet(ref tempNwireCustomLevelSheet, singleLevelSheetData.Value);
                levelSheet.Rows.AddRange(tempNwireCustomLevelSheet.Rows);
                if (levelSheet.Name.StartsWith("Levels_Rtos"))
                {
                    _rtosLevelsGenerator.UpdateLevelSheet(ref levelSheet);
                }

                levelSheets.Add(levelSheet);
            }
            return levelSheets;
        }

        #region Create HardIpDc Level Sheet

        /// <summary>
        /// Create special level sheets using information from HardIP DC sheet
        /// </summary>
        /// <param name="hardIpDc">Hardip dc sheet</param>
        /// <param name="ioContiSheet"></param>
        /// <returns>special level sheets</returns>
        protected List<LevelSheet> CreateHardIpDcLevels(LevelSheet hardIpLevel, HardIpDcSheet hardIpDc, IoContiSheet ioContiSheet)
        {
            List<LevelSheet> hardIpLevelSheets = new List<LevelSheet>();
            if (hardIpDc == null)
            {
                return hardIpLevelSheets;
            }
            try
            {
                foreach (HardIpCategoryDef catDef in hardIpDc.Rows)
                {
                    if (catDef.LevelSheet.Equals("Levels_HardIP", StringComparison.OrdinalIgnoreCase))
                    {
                        //when only Power pins and Levels I/O pins specified in HardIp Dc sheet, No need to Create a new Level sheet 
                        continue;
                    }
                    LevelSheet levelSheet = hardIpLevel.Copy();
                    levelSheet.Name = catDef.LevelSheet;
                    AddLevelRowFromHardIpDc(levelSheet, catDef, ioContiSheet);
                    hardIpLevelSheets.Add(levelSheet);
                }
            }
            catch (Exception)
            {
                //pass
            }
            return hardIpLevelSheets;
        }

        /// <summary>
        /// Add Level rows into Level sheet from HardIp Dc sheet
        /// </summary>
        /// <param name="level">Default level sheet</param>
        /// <param name="categoryDef">HardIP Category definition</param>
        /// <param name="ioContiSheet"></param>
        public virtual void AddLevelRowFromHardIpDc(LevelSheet level, HardIpCategoryDef categoryDef, IoContiSheet ioContiSheet = null)
        {
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;

            foreach (HardIpDcRow row in categoryDef.DataRows)
            {
                List<HardIpSpecValue> specValues = categoryDef.GetSpecValueFromDef(row);
                foreach (HardIpSpecValue specValue in specValues)
                {
                    if (specValue.HasRatio)
                    {
                        //These specs will defined in DC Spec, not level sheet
                        continue;
                    }

                    LevelRow levelRow = level.Rows.Find(p => p.PinName.Equals(specValue.PinName, StringComparison.OrdinalIgnoreCase)
                             && p.Parameter.Equals(specValue.Parameter, StringComparison.OrdinalIgnoreCase));

                    if (levelRow != null && !specValue.FromDefault)
                    {
                        //Power pin or like Pins_1p1v and specify in HardIP_DC sheet, Modify exist value
                        levelRow.Value = specValue.Value;
                    }
                    else if (levelRow != null && specValue.FromDefault)
                    {
                        //Power pin or like Pins_1p1v and not specify in HardIP_DC sheet, No need to modify exist value
                    }
                    else if (levelRow == null && specValue.FromDefault)
                    {
                        #region no level & check io conti
                        string ioGrpName = GetPinGroupOfNonRatioPin(specValue, ioContiSheet);
                        if (ioGrpName != "")  //if value fromDefault, 
                        {
                            LevelRow oneRow = level.Rows.Find(p => p.PinName.Equals(ioGrpName, StringComparison.OrdinalIgnoreCase) &&
                                p.Parameter.Equals(specValue.Parameter, StringComparison.OrdinalIgnoreCase));
                            if (oneRow != null)
                            {
                                specValue.Value = oneRow.Value;
                            }
                        }
                        else
                        {
                            if (pinMap != null)
                            {
                                if (pinMap.IsGroupExist(row.PinName))
                                {
                                    List<string> voltageList = new List<string>();
                                    List<Pin> list = pinMap.GetPinsFromGroup(row.PinName);
                                    foreach (Pin item in list)
                                    {
                                        if (ioContiSheet != null)
                                        {
                                            voltageList.Add(ioContiSheet.GetVoltage(item.PinName));
                                        }
                                    }
                                    if (voltageList.Distinct().Count() == 1)
                                    {
                                        string voltage = voltageList.Distinct().First();
                                        string pinGrpName = "Pins_" + voltage.Replace(".", "p") + "v";
                                        pinGrpName = pinGrpName.Replace("0v", "v");
                                        LevelRow oneRow = level.Rows.Find(p => p.PinName.Equals(pinGrpName, StringComparison.OrdinalIgnoreCase) &&
                                            p.Parameter.Equals(specValue.Parameter, StringComparison.OrdinalIgnoreCase));
                                        if (oneRow != null)
                                        {
                                            specValue.Value = oneRow.Value;
                                        }
                                    }
                                    else
                                    {
                                        string outString = "The voltage of IO pin group " + row.PinName + " had different voltage";
                                        ErrorReportManager.AddError(BasicErrorType.E_MissingPin_19, categoryDef.LevelSheet, 1, 0, "The voltage of IO pin group " + row.PinName + " had different voltage", new string[] { row.PinName }, new ErrorInfo() { Comments = new List<string>() { row.PinName } });
                                    }
                                }
                            }
                        }

                        //Other I/O pins, Append to Level sheet
                        level.AddRow(specValue);
                        #endregion
                    }
                    else
                    {
                        //Other I/O pins, Append to Level sheet
                        level.AddRow(specValue);
                    }
                }
            }
        }

        public string GetPinGroupOfNonRatioPin(HardIpSpecValue specValue, IoContiSheet ioContiSheet)
        {
            string retStr = "";
            if (ioContiSheet == null)
            {
                return retStr;
            }

            const string regIoVoltage = @"(?<num>\d*)\.(?<digi>\d*)";
            string ioVoltage = ioContiSheet.GetVoltage(specValue.PinName);
            if (ioVoltage != "")
            {
                string volStr = ioVoltage;  //ex: 1.8V
                string numTok = Regex.Match(volStr, regIoVoltage).Groups["num"].ToString();
                string digTok = Regex.Match(volStr, regIoVoltage).Groups["digi"].ToString();
                retStr = $"Pins_{numTok}p{digTok}v";  //Pins_1p8v
            }
            return retStr;
        }
        #endregion

        #region Create Nwire Level sheet

        protected LevelSheet CreateNwireLevelSheet()
        {
            string sheetName = "Levels_nWire";
            var nWireLevel = new LevelSheet(sheetName);
            _nwireLevelsGenerator.UpdateLevelSheet(ref nWireLevel);
            return nWireLevel;
        }

        #endregion

        #region Create Spi rom Level sheet
        /// <summary>
        /// Create Level sheet for SpiRom
        /// </summary>
        /// <returns></returns>
        internal LevelSheet CreateSpiRomLevelSheet()
        {
            string sheetName = "Levels_Write_SPIROM";
            LevelSheet spiromLevel = new LevelSheet(sheetName);
            _nwireLevelsGenerator.UpdateLevelSheet(ref spiromLevel);
            return spiromLevel;
        }
        #endregion

        public virtual LevelSheet CreateLevelSheet(string levelSheetName, bool isGenVdd = true, string suffix = "", string chiplet = "", PowerInfoSheet powerinfo = null, bool isHardipDc = false)
        {
            if (!string.IsNullOrEmpty(chiplet))
            {
                levelSheetName = levelSheetName + "_" + chiplet;
            }
            var levelSheet = new LevelSheet(levelSheetName);

            return levelSheet;
        }
    }
}
