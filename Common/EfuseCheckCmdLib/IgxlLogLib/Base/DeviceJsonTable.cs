using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;

using LogLib.Utility;

using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class DeviceJsonTable
    {
        public string JsonTableDirectory = Path.Combine(Directory.GetCurrentDirectory(), "ValidationTemp", "DeviceJson");

        public ConcurrentDictionary<string, string> DeviceJsonPathDict = new();
        private BlockingCollection<DeviceData> _queueTable = null!;

        public bool SkipParsing { get; private set; }

        public List<string> GetDeviceList
        {
            get { return [.. DeviceJsonPathDict.Keys]; }
        }

        public DeviceJsonTable(string subFolder, string deviceFileBaseName, int deviceCnt)
        {

            JsonTableDirectory = Path.Combine(JsonTableDirectory, subFolder, deviceFileBaseName);

            if (Directory.Exists(JsonTableDirectory))
            {
                if (Directory.GetFiles(JsonTableDirectory).Length == deviceCnt)
                {
                    SkipParsing = true;
                    return;
                }

                Directory.Delete(JsonTableDirectory, true);
            }

            Directory.CreateDirectory(JsonTableDirectory);

        }

        public void InitialQueue()
        {
            _queueTable = [];
        }

        public void SetDeviceDataToQueue(DeviceData deviceData)
        {
            _queueTable.Add(deviceData);
        }

        public void QueueComplete()
        {
            _queueTable.CompleteAdding();
        }

        public void UpdateDataBaseByQueue()
        {
            try
            {
                foreach (DeviceData deviceData in _queueTable.GetConsumingEnumerable())
                {
                    //DumpDatableToJson(table);
                    DumpDeviceDataToJson(deviceData);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        public DataTable GetDataTable<T>(string device) where T : ILogRow
        {

            List<T> data = GetDeserializeObjList<T>(device);

            IContractResolver resolver = new DefaultContractResolver();
            // to get json property Name

            string[] propertyNames = resolver.PropertyNames(typeof(T));

            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            var dt = new DataTable();

            for (int i = 0; i < propertyNames.Length; i++)
            {
                PropertyDescriptor prop = props[i];
                dt.Columns.Add(propertyNames[i], prop.PropertyType);
            }

            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item)!;
                }
                dt.Rows.Add(values);
            }

            return dt;
        }

        public Queue<T> GetDeserializeObjQueue<T>(string device)
        {
            string deviceJsonFile = DeviceJsonPathDict[device];
            Queue<T> jsonLoad = new Queue<T>();
            using (var r = new StreamReader(deviceJsonFile))
            {
                var jsonserializer = new JsonSerializer();
                using var jsonReader = new JsonTextReader(r);

                while (jsonReader.Read())
                {
                    if (jsonReader.TokenType == JsonToken.StartObject)
                    {
                        jsonLoad.Enqueue(jsonserializer.Deserialize<T>(jsonReader)!);
                    }
                }
            }
            return jsonLoad;
        }

        public List<T> GetDeserializeObjList<T>(string device)
        {
            string deviceJsonFile = DeviceJsonPathDict[device];
            var jsonLoad = new List<T>();
            using (var r = new StreamReader(deviceJsonFile))
            {
                var jsonserializer = new JsonSerializer();
                using var jsonReader = new JsonTextReader(r);

                while (jsonReader.Read())
                {
                    if (jsonReader.TokenType == JsonToken.StartObject)
                    {
                        jsonLoad.Add(jsonserializer.Deserialize<T>(jsonReader)!);
                    }
                }
            }
            return jsonLoad;
        }

        private void DumpDeviceDataToJson(DeviceData deviceData)
        {
            string jsonFile = Path.Combine(JsonTableDirectory, deviceData.DeviceNum + ".txt");

            DeviceJsonPathDict[deviceData.DeviceNum.ToString()] = jsonFile;

            if (!File.Exists(jsonFile))
            {
                using StreamWriter file = File.CreateText(jsonFile);
                JsonSerializer siSerializer = new JsonSerializer { Formatting = Formatting.Indented };
                siSerializer.Serialize(file, deviceData.DataLowRows);
            }
        }

        //private void DumpDatableToJson(DataTable dt)
        //{
        //    var jsonFile = Path.Combine(JsonTableDirectory, dt.TableName + ".txt");

        //    DeviceJsonPathDict[dt.TableName] = jsonFile;

        //    if (!File.Exists(jsonFile))
        //    {
        //        using (StreamWriter file = File.CreateText(jsonFile))
        //        {
        //            JsonSerializer siSerializer = new JsonSerializer();
        //            siSerializer.Formatting = Formatting.Indented;
        //            siSerializer.Serialize(file, dt);
        //        }
        //    }
        //}
    }
}
