using Automation.Static;

using CommonLib.Enums;

namespace Automation.Library
{
    public class ClassVbt
    {
        public static List<string> ReadBasContent(string filename)
        {
            var content = new List<string>();
            var reader = new StreamReader(filename);
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                content.Add(line);
            }
            reader.Close();
            return content;
        }

        public static string SearchContent(List<string> content, List<string> patterns)
        {
            bool hasSearch = true;
            foreach (string perLine in content)
            {
                foreach (string pattern in patterns)
                {
                    if (!perLine.ToLower().Contains(pattern.ToLower()))
                    {
                        hasSearch = false;
                        break;
                    }
                    hasSearch = true;
                }
                if (hasSearch)
                {
                    return perLine;
                }
            }
            return "";
        }

        public static void WriteVBFile(string filename, List<string> content)
        {
            var writer = new StreamWriter(filename);
            foreach (string perLine in content)
            {
                writer.WriteLine(perLine);
            }
            writer.Close();
        }

        public static void ModifyCommonBas(List<FileInfo> mFileList)
        {
            if (!mFileList.Any())
            {
                return;
            }

            FileInfo targetBas =
                mFileList.Find(p => p.Name.Equals("LIB_Common_GlobalConstant.bas", StringComparison.OrdinalIgnoreCase));
            List<string> content = ReadBasContent(targetBas.FullName);
            int dcviIndex =
                content.FindIndex(
                    x => x.Equals(SearchContent(content, new List<string> { "AllDCVIPinlist", "Const" })));
            int powerIndex =
                content.FindIndex(
                    p => p.Equals(SearchContent(content, new List<string> { "AllPowerPinlist", "Const" })));
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                if (!content[dcviIndex].Split(' ').Last().Equals("\"DCVI_Power\"", StringComparison.OrdinalIgnoreCase))
                {
                    content[dcviIndex] = "Public Const AllDCVIPinlist = \"DCVI_Power\"";
                }
                if (!content[powerIndex].Split(' ').Last().Equals("\"DCVS_Power\"", StringComparison.OrdinalIgnoreCase))
                {
                    content[powerIndex] = "Public Const AllPowerPinlist = \"DCVS_Power\"";
                }
            }
            else
            {
                if (!content[dcviIndex].Split(' ').Last().Equals("\"All_DCVI\"", StringComparison.OrdinalIgnoreCase))
                {
                    content[dcviIndex] = "Public Const AllDCVIPinlist = \"All_DCVI\"";
                }
                if (!content[powerIndex].Split(' ').Last().Equals("\"All_Power\"", StringComparison.OrdinalIgnoreCase))
                {
                    content[powerIndex] = "Public Const AllDCVIPinlist = \"All_Power\"";
                }
            }
            WriteVBFile(targetBas.FullName, content);
        }

        public static void ModifyExecIpBas(List<FileInfo> mFileList)
        {
            FileInfo targetBas = mFileList.Find(p => p.Name.Equals("Exec_IP_Module.bas", StringComparison.OrdinalIgnoreCase));
            if (targetBas == null)
            {
                return;
            }

            List<string> content = ReadBasContent(targetBas.FullName);
            int index =
                content.FindIndex(
                    x => x.Equals(SearchContent(content, new List<string> { "Exec_IP_Module", "Attribute", "VB_Name" })));
            if (index != -1)
            {
                if (!content.Exists(p => p.Contains("#Const AP =")))
                {
                    content.Insert(index + 1, $"#Const AP = {LocalSpecs.Options.Device == EnumDevice.AP}");
                }
                if (!content.Exists(p => p.Contains("#Const RF =")))
                {
                    content.Insert(index + 1, $"#Const RF = {LocalSpecs.Options.Device == EnumDevice.RF}");
                }
                if (!content.Exists(p => p.Contains("#Const LCD =")))
                {
                    content.Insert(index + 1, $"#Const LCD = {LocalSpecs.Options.Device == EnumDevice.LCD}");
                }
                //if (ProjectConfigLocalSpecs.Device == DeviceEnum.AP)
                //{
                //    content.Insert(index + 1, "#Const AP = True");
                //    content.Insert(index + 2, "#Const RF = False");
                //    content.Insert(index + 3, "#Const LCD = False");
                //}
                //else if (ProjectConfigLocalSpecs.Device == DeviceEnum.RF)
                //{
                //    content.Insert(index + 1, "#Const AP = False");
                //    content.Insert(index + 2, "#Const RF = True");
                //    content.Insert(index + 3, "#Const LCD = False");
                //}
                //else if (ProjectConfigLocalSpecs.Device == DeviceEnum.LCD)
                //{
                //    content.Insert(index + 1, "#Const AP = False");
                //    content.Insert(index + 2, "#Const RF = False");
                //    content.Insert(index + 3, "#Const LCD = True");
                //}
            }
            WriteVBFile(targetBas.FullName, content);
        }
    }
}
