using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

namespace ProjectConfigLib.ProjectConfig
{
    public class ProjectConfigSingleton
    {
        private static ProjectConfigSingleton _instance = null!;
        private List<ProjectConfigRow> _projectConfigRows = [];
        private List<ProjectConfigSettingRow> _projectConfigSettingRows = null!;

        private ProjectConfigSingleton()
        {
            InitializeProjectConfigSetting();
        }

        public static ProjectConfigSingleton Instance()
        {
            return _instance ??= new ProjectConfigSingleton();
        }

        public void InitializeProjectConfigSetting(string projectName = "")
        {
            _projectConfigSettingRows = ProjectConfigSettingReader.Read(projectName)!;
        }

        public List<ProjectConfigSettingRow> GetRows()
        {
            return _projectConfigSettingRows;
        }

        public bool GetBool(string groupName, string name)
        {
            string value = GetValue(groupName, name);
            return value.EqualsIgnoreCase("TRUE");
        }

        public int GetInt(string groupName, string name, int defaultValue = 0)
        {
            string value = GetValue(groupName, name);

            return int.TryParse(value, out int result) ? result : defaultValue;
        }

        public string GetValue(string groupName, string name)
        {
            var value = _projectConfigRows.Where(t => t.Group.EqualsIgnoreCase(groupName) && t.Name.EqualsIgnoreCase(name)).ToList();
            if (value.Count > 0)
            {
                return value[0].Value;
            }

            ProjectConfigSettingRow? defaultSetting = _projectConfigSettingRows?.Find(t => t.Name.EqualsIgnoreCase(name));
            return defaultSetting != null ? defaultSetting.Default : "";
        }

        public void LoadProjectConfig(string iniFileName)
        {
            _projectConfigRows.Clear();

            if (!File.Exists(iniFileName))
            {
                //Get from default
                _projectConfigRows = [.. GetRows().Select(x => x.ConvertToProjectConfigRow())];
            }
            else
            {
                //Get read from *.ini by _comIni
                _projectConfigRows = ProjectConfigIniReader.ReadFile(iniFileName);
            }
        }
    }
}
