using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.InputReader
{
    public class ReadFlowSheet
    {

        #region New
        /// <summary>
        /// ReadFlowSheet
        /// </summary>
        /// <param name="pFlowSheet"></param>
        public ReadFlowSheet()
        {

        }

        private static SheetObjMap GetSheetFormat(IIgxlSheet sheet)
        {
            return LocalSpecs.IgxlConfig.SheetItemClass.FirstOrDefault(p => p.sheetname.Equals(sheet.IgxlSheetName));
        }
        #endregion

        #region public Function
        /// <summary>
        /// GetSheet
        /// </summary>
        /// <returns></returns>
        public SubFlowSheet GetSheet(string path)
        {
            SubFlowSheet lResultIgxlSheet = new SubFlowSheet(Path.GetFileNameWithoutExtension(path));
            List<ClassProperty> format = GetSheetFormat(lResultIgxlSheet).InnerObj[0].Property.ToList();
            var reader = new StreamReader(path);
            string line = "";
            var header = new Dictionary<int, string>();

            while ((line = reader.ReadLine()) != null)
            {
                List<string> rowItems = line.Split('\t').ToList();
                if (header.Count == 0)
                {
                    if (format.Exists(p => rowItems.Contains(p.nameInSheet)))
                    {
                        for (int i = 0; i < format.Count; i++)
                        {
                            header.Add(format[i].indexInSheet, format[i].name);
                        }
                    }
                }
                else
                {
                    FlowRow row = ParseFlowRow(rowItems, header);
                    if (string.IsNullOrEmpty(row.Opcode))
                    {
                        break;
                    }
                    lResultIgxlSheet.AddRow(row);
                }
            }
            reader.Close();
            return lResultIgxlSheet;
        }
        #endregion

        //Label	Enable	Job	Part	Env
        private static FlowRow ParseFlowRow(List<string> rowItems, Dictionary<int, string> header)
        {
            var row = new FlowRow();
            for (int i = 0; i < rowItems.Count; i++)
            {
                string rowItem = rowItems[i];
                var keys = header.Select(p => p.Value).ToList();
                string xxx = string.Join("\t", keys);
                if (header.TryGetValue(i, out string value))
                {
                    AssignFlowRowField(row, value, rowItem);
                }
                //Elapsed Time (s)	Background Type	Serialize	Resource Lock	Flow Step Locked	Comment
            }
            return row;
        }

        private static void AssignFlowRowField(FlowRow row, string fieldKey, string rowItem)
        {
            if (AssignCoreField(row, fieldKey, rowItem))
            {
                return;
            }
            if (AssignBinResultField(row, fieldKey, rowItem))
            {
                return;
            }
            AssignGroupDeviceDebugField(row, fieldKey, rowItem);
        }

        private static bool AssignCoreField(FlowRow row, string fieldKey, string rowItem)
        {
            switch (fieldKey)
            {
                case "Label":
                    row.Label = rowItem;
                    return true;
                case "EnableField":
                    row.Enable = rowItem;
                    return true;
                case "GateJob":
                    row.Job = rowItem;
                    return true;
                case "GatePart":
                    row.Part = rowItem;
                    return true;
                case "GateEnv":
                    row.Env = rowItem;
                    return true;
                case "Opcode":
                    row.Opcode = rowItem;
                    return true;
                case "Parameter":
                    row.Parameter = rowItem;
                    return true;
                case "TName":
                    row.TName = rowItem;
                    return true;
                case "TNum":
                    row.TNum = rowItem;
                    return true;
                case "LoLim":
                    row.LoLim = rowItem;
                    return true;
                case "HiLim":
                    row.HiLim = rowItem;
                    return true;
                case "DatalogScale":
                    row.Scale = rowItem;
                    return true;
                case "DatalogUnits":
                    row.Units = rowItem;
                    return true;
                case "DatalogFormat":
                    row.Format = rowItem;
                    return true;
                default:
                    return false;
            }
        }

        private static bool AssignBinResultField(FlowRow row, string fieldKey, string rowItem)
        {
            switch (fieldKey)
            {
                case "HardBinPass":
                    row.BinPass = rowItem;
                    return true;
                case "HardBinFail":
                    row.BinFail = rowItem;
                    return true;
                case "SoftBinPass":
                    row.SortPass = rowItem;
                    return true;
                case "SoftBinFail":
                    row.SortFail = rowItem;
                    return true;
                case "Result":
                    row.Result = rowItem;
                    return true;
                case "PassAction":
                    row.PassAction = rowItem;
                    return true;
                case "FailAction":
                    row.FailAction = rowItem;
                    return true;
                case "State":
                    row.State = rowItem;
                    return true;
                case "Comment":
                    row.Comment = rowItem;
                    return true;
                default:
                    return false;
            }
        }

        private static void AssignGroupDeviceDebugField(FlowRow row, string fieldKey, string rowItem)
        {
            switch (fieldKey)
            {
                case "GroupSpecifier":
                    row.GroupSpecifier = rowItem;
                    break;
                case "GroupSense":
                    row.GroupSense = rowItem;
                    break;
                case "GroupCondition":
                    row.GroupCondition = rowItem;
                    break;
                case "GroupName":
                    row.GroupName = rowItem;
                    break;
                case "DeviceSense":
                    row.DeviceSense = rowItem;
                    break;
                case "DeviceCondition":
                    row.DeviceCondition = rowItem;
                    break;
                case "DeviceName":
                    row.DeviceName = rowItem;
                    break;
                case "DebugAssume":
                    row.DebugAssume = rowItem;
                    break;
                case "DebugSites":
                    row.DebugSites = rowItem;
                    break;
                case "CtProfileDataElapsedTime":
                    row.CtProfileDataElapsedTime = rowItem;
                    break;
                case "CtProfileDataBackgroundType":
                    row.CtProfileDataBackgroundType = rowItem;
                    break;
                case "CtProfileDataSerialize":
                    row.CtProfileDataSerialize = rowItem;
                    break;
                case "CtProfileDataResourceLock":
                    row.CtProfileDataResourceLock = rowItem;
                    break;
                case "CtProfileDataFlowStepLocked":
                    row.CtProfileDataFlowStepLocked = rowItem;
                    break;
                default:
                    break;
            }
        }

    }
}
