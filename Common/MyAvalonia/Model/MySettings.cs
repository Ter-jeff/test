using System.Collections.Generic;

namespace MyAvalonia.Model
{
    public class MySettings
    {
        public List<IMySetting> Collection { get; set; } = [];

        public void Write(string fileName, IMySetting setting)
        {
            string classType = setting.ClassType;
            if (Collection.Exists(x => x.ClassType == classType))
            {
                Collection.RemoveAll(x => x.ClassType == classType);
            }
            Collection.Add(setting);
            Json.Write(fileName, this);
        }
    }
}
