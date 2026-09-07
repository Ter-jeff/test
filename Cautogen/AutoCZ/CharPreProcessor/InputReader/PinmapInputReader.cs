using System.Collections.Generic;
using System.IO;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader;

using CommonLib.Enums;

using LogLib.Utility;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    public class PinmapInputReader : TextInputReader
    {
        // constructor
        public PinmapInputReader(string filePath)
            : base(filePath) { }

        protected override void _Read(StreamReader textReader)
        {
            MessageWriter.WriteMessage("Reading PinMap file...", EnumMessageLevel.Info);
            LogHelper.Info("Reading PinMap file...");

            string line = "";
            while (line != null && !line.Contains("Pin Name"))
            {
                line = textReader.ReadLine();
            }

            while ((line = textReader.ReadLine()) != null)
            {
                if (line == "")
                {
                    continue;
                }

                string key;
                string value, type;

                if (line.Split('\t')[1] == "")
                {
                    value = line.Split('\t')[2];
                    type = line.Split('\t')[3];
                    key = value.Replace("_", "").ToLower();
                }
                else
                {
                    value = line.Split('\t')[1];
                    key = value.Replace("_", "").ToLower();
                    string pin = line.Split('\t')[2].ToUpper();
                    type = string.IsNullOrEmpty(line.Split('\t')[3]) ? "" : line.Split('\t')[3];

                    if (!UtilityMain.UtilityData.PinGroups.Keys.Contains(value.ToUpper()))
                    {
                        var pinList = new List<string> { pin };
                        UtilityMain.UtilityData.PinGroups.Add(value.ToUpper(), pinList);
                        UtilityMain.UtilityData.PinGroupsType.Add(value.ToUpper(), type);
                    }
                    else
                    {
                        UtilityMain.UtilityData.PinGroups[value.ToUpper()].Add(pin);
                        if (UtilityMain.UtilityData.PinGroupsType[value.ToUpper()] == "")
                        {
                            UtilityMain.UtilityData.PinGroupsType[value.ToUpper()] = type;
                        }
                    }
                }
                if (!UtilityMain.UtilityData.PinList.Keys.Contains(key))
                {
                    UtilityMain.UtilityData.PinList.Add(key, new PinInfo(value, type));
                }
            }
        }
    }
}
