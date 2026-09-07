using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.common.ReaderWriter.Reader;

using Newtonsoft.Json.Linq;

namespace Cautogen.AutoCZ.CharPostProcessor.InputReader
{
    public class PostSettings : TextInputReader
    {
        /* field */
        public static List<string> PreservedSheets = new List<string>();

        /* constructor */
        public PostSettings(string filePath)
            : base(filePath) { }

        /* method */
        protected override void _Read(StreamReader textReader)
        {
            // reset previous result
            PreservedSheets = new List<string>();

            // parse settings            
            try
            {
                string allText = textReader.ReadToEnd();
                var settings = JObject.Parse(allText);
                foreach (string sheetStr in settings["preserve_sheets"]
                    .Select(sheet => sheet.ToString())
                    .Select(sh => Regex.IsMatch(@"\.txt$", sh, RegexOptions.IgnoreCase) ? sh : sh + ".txt"))
                {
                    PreservedSheets.Add(sheetStr);
                }
            }
            catch (Exception e)
            {
                LocalSpecs.MessageWriter.WriteLine(e.Message);
            }
        }
    }
}
