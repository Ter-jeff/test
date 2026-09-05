using System;
using System.Collections.Generic;
using System.IO;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader;

using CommonLib.Enums;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    public class SelsramTableReader : TextInputReader
    {
        /* properties */
        public static List<SelSrmRow> SelSramRows = new List<SelSrmRow>();

        /* constructor */
        public SelsramTableReader(string filePath)
            : base(filePath)
        {
        }

        /* methods */
        protected override void _Read(StreamReader textReader)
        {
            MessageWriter.WriteMessage("Reading Selsram Mapping Table...", EnumMessageLevel.Info);
            SelSramRows.Clear();
            int i = 0;
            var header = new List<string>();
            string line;
            while ((line = textReader.ReadLine()) != null && !line.ToLower().Contains("end"))
            {
                //if (line.Contains("VDD_"))
                {
                    string[] values = line.Split(new[] { '\t' });
                    if (i == 0)
                    {
                        header.AddRange(line.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries));
                        i++;
                        continue;
                    }
                    SelSrmRow row = new SelSrmRow
                    {
                        RowNum = i, Bits = values[header.IndexOf("Bits")], LogicPins = values[header.IndexOf("Logic Pins")],
                        SramPins = values[header.IndexOf("Sram Pins")],
                        Selsrm1 = values[header.IndexOf("Selsrm1")],
                        Selsrm0 = values[header.IndexOf("Selsrm0")]
                    };

                    SelSramRows.Add(row);
                    i++;
                }
            }

        }


    }
}
