using System;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace CommonLib.Extension
{
    public static class Clone
    {
        public static T DeepClone<T>(this T source)
        {
            if (!typeof(T).IsSerializable)
            {
                throw new ArgumentException("The type must be serializable.");
            }

            if (source != null)
            {
                using (var stream = new MemoryStream())
                {
                    var formatter = new BinaryFormatter();
                    formatter.Serialize(stream, source);
                    stream.Seek(0, SeekOrigin.Begin);
                    var clonedSource = (T)formatter.Deserialize(stream);
                    return clonedSource;
                }
            }

            return default(T);
        }

        //public static T CloneJson<T>(this T source)
        //{
        //    // Don't serialize a null object, simply return the default for that object
        //    if (ReferenceEquals(source, null)) return default(T);

        //    // initialize inner objects individually
        //    // for example in default constructor some list property initialized with some values,
        //    // but in 'source' these items are cleaned -
        //    // without ObjectCreationHandling.Replace default constructor values will be added to result
        //    var deserializeSettings = new JsonSerializerSettings
        //    { ObjectCreationHandling = ObjectCreationHandling.Replace };
        //    return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(source), deserializeSettings);
        //}
    }
}
