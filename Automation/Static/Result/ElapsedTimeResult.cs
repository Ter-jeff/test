using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace Automation.Static.Result
{
    public class ElapsedTimeResults : List<ElapsedTimeResult>
    {
        public void MeasureSection(string name, Action action)
        {
            var sw = Stopwatch.StartNew();
            action();
            sw.Stop();
            TimeSpan elapsed = sw.Elapsed;
            string text = elapsed.ToString(@"hh\:mm\:ss\.fff");
            if (TimeSpan.TryParse(text, out TimeSpan ts))
            {
                double sec = ts.TotalSeconds;
                this.Add(new ElapsedTimeResult(name, sec));
            }
        }
        public void WriteToJson()
        {
            if (!this.Any())
            {
                return;
            }

            var serializer = new DataContractJsonSerializer(GetType());
            string outPutFile = Path.Combine(FolderStructure.TarDir, "ElapsedTime.json");
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, this);
                stream.Position = 0;
                using (var sr = new StreamReader(stream))
                {
                    using (var sw = new StreamWriter(outPutFile))
                    {
                        sw.WriteLine(sr.ReadToEnd());
                        sw.Close();
                    }
                    sr.Close();
                }
                stream.Close();
            }
        }
    }

    [DataContract]
    public class ElapsedTimeResult
    {
        [DataMember]
        public string Module { get; set; }
        [DataMember]
        public double Time { get; set; }
        public ElapsedTimeResult(string module, double time)
        {
            Module = module;
            Time = time;
        }
    }
}
