using System;
using System.IO;

using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace MyAvalonia.Model
{
    public static class Json
    {
        public static void Write<T>(string fileName, T item)
        {
            string jsonData = JsonConvert.SerializeObject(item, Formatting.Indented, new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto, SerializationBinder = new TypeNameSerializationBinder() });
            string? dir = Path.GetDirectoryName(fileName);
            if (!Directory.Exists(dir) && dir != null)
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(fileName, jsonData);
        }

        public static T? Read<T>(string fileName)
        {
            string text = File.ReadAllText(fileName);
            return JsonConvert.DeserializeObject<T>(text, new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto, SerializationBinder = new TypeNameSerializationBinder() });
        }

        public class TypeNameSerializationBinder : ISerializationBinder
        {
            public void BindToName(Type serializedType, out string? assemblyName, out string? typeName)
            {
                assemblyName = serializedType.Assembly.GetName().Name;
                typeName = serializedType.FullName;
            }

            public Type BindToType(string? assemblyName, string? typeName)
            {
                foreach (System.Reflection.Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    foreach (Type type in assembly.GetTypes())
                    {
                        if (type.FullName == typeName)
                        {
                            return type;
                        }
                    }
                }
                throw new InvalidOperationException($"Cannot find type '{typeName}'");
            }
        }
    }
}
