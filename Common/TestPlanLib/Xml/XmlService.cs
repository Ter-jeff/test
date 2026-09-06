using System;
using System.IO;
using System.Xml.Serialization;

namespace TestPlanLib.Xml
{
    public static class XmlService<T>
    {
        public static T LoadXml(string fileName)
        {
            T result;
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(T));
                StreamReader sr = new StreamReader(fileName);
                T sysData = (T)xs.Deserialize(sr)!;
                sr.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception(e.StackTrace);
            }
            return result;
        }
    }
}
