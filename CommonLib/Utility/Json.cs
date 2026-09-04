using System;
using System.IO;

using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace CommonLib.Utility
{
    public static class Json
    {
        public static void Write<T>(string fileName, T item)
        {
            string jsonData = JsonConvert.SerializeObject(item, Formatting.Indented, new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto, SerializationBinder = new TypeNameSerializationBinder() });
            string dir = Path.GetDirectoryName(fileName);
            if (!Directory.Exists(dir))
            {
                if (dir != null)
                {
                    Directory.CreateDirectory(dir);
                }
            }
            File.WriteAllText(fileName, jsonData);
        }

        public static T Read<T>(string fileName)
        {
            string text = File.ReadAllText(fileName);
            T descJsonStu = JsonConvert.DeserializeObject<T>(text, new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto, SerializationBinder = new TypeNameSerializationBinder() });
            return descJsonStu;
        }

        public class TypeNameSerializationBinder : ISerializationBinder
        {
            public void BindToName(Type serializedType, out string assemblyName, out string typeName)
            {
                //assemblyName = "TeradyneConfig";
                assemblyName = serializedType.Assembly.GetName().Name;
                typeName = serializedType.FullName;
            }

            public Type BindToType(string assemblyName, string typeName)
            {
                System.Reflection.Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
                foreach (System.Reflection.Assembly assembly in assemblies)
                {
                    Type[] types = assembly.GetTypes();
                    foreach (Type type in types)
                    {
                        if (type.FullName == typeName)
                        {
                            return type;
                        }
                    }
                }
                throw new InvalidOperationException();
            }
        }
    }
}
