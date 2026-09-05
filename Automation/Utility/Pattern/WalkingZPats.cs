using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Automation.Utility.Pattern
{
    public class WalkingZPats
    {
        public string OutputFolder;
        public int Repeat;

        public WalkingZPats(string outputFolder, int repeat)
        {
            OutputFolder = outputFolder;
            if (!Directory.Exists(OutputFolder))
            {
                Directory.CreateDirectory(OutputFolder);
            }

            Repeat = repeat;
        }

        public void WorkFlowFromContiSheet(List<string> allIOs, List<string> testIOs, string moduleName, Dictionary<string, string> diffPairs)
        {
            WritePatterns(allIOs, testIOs, moduleName, diffPairs);
        }

        private void WritePatterns(List<string> allIOs, List<string> testIOs, string moduleName, Dictionary<string, string> diffPairs)
        {
            var allIOsClone = allIOs.Select(a => a).ToList();
            var sw = new StreamWriter(Path.Combine(OutputFolder, moduleName + ".atp"));
            sw.WriteLine("digital_inst = hsdmq;");
            sw.WriteLine("opcode_mode = single;");
            sw.WriteLine("import tset Tset;");
            int maxPinLen = testIOs.Max(a => a.Length);
            WriteHeader(sw, allIOsClone, testIOs, moduleName);
            WritePinsInVertical(sw, allIOsClone, testIOs, maxPinLen);

            #region print patterns
            sw.WriteLine("\n");
            sw.WriteLine(moduleName + ":");
            int testCnt = testIOs.Count;
            WriteWalkingRows(sw, allIOsClone, testIOs, diffPairs, testCnt);
            string pinLine;
            for (int iX = 0; iX < 64 - testCnt; iX++)
            {
                pinLine = $"{"> Tset ",21}";
                for (int iY = 0; iY < testCnt; iY++)
                {
                    pinLine += " X";
                }

                if (allIOsClone.Count > 0)
                {
                    pinLine += " .X;";
                }
                else
                {
                    pinLine += " ;";
                }

                sw.WriteLine(pinLine);
            }
            pinLine = $"{"HALT > Tset ",21}";
            for (int iY = 0; iY < testCnt; iY++)
            {
                pinLine += " X";
            }

            if (allIOsClone.Count > 0)
            {
                pinLine += " .X;\n}";
            }
            else
            {
                pinLine += " ;\n}";
            }

            sw.WriteLine(pinLine);
            #endregion

            sw.Close();
        }

        private void WritePinsInVertical(StreamWriter sw, List<string> allIOsClone, List<string> testIOs, int maxPinLen)
        {
            int iLen;
            #region print pin in vertical
            string pinLine;
            for (iLen = 0; iLen < maxPinLen; iLen++)
            {
                pinLine = $"{"",-19}/*";
                foreach (string pin in testIOs)
                {
                    string s;
                    if (pin.Length > iLen)
                    {
                        s = " " + pin.Substring(iLen, 1);
                    }
                    else
                    {
                        s = "  ";
                    }
                    pinLine += s;
                }
                if (allIOsClone.Count > 0)
                {
                    pinLine += " Others */";
                }
                else
                {
                    pinLine += " */";
                }

                sw.WriteLine(pinLine);
            }
            #endregion
        }

        private void WriteWalkingRows(StreamWriter sw, List<string> allIOsClone, List<string> testIOs, Dictionary<string, string> diffPairs, int testCnt)
        {
            string pinLine;
            for (int iX = 0; iX < testCnt; iX++)
            {
                int diffCol = -1;
                if (diffPairs.ContainsKey(testIOs[iX]))
                {
                    string pair = diffPairs[testIOs[iX]];
                    diffCol = testIOs.FindIndex(x => x == pair);
                }

                if (Repeat > 0)
                {
                    pinLine = $"repeat {Repeat,-4} {"> Tset ",9}";
                    for (int iY = 0; iY < testCnt; iY++)
                    {
                        if (iX == iY || diffCol == iY)
                        {
                            pinLine += " X";
                        }
                        else
                        {
                            pinLine += " 0";
                        }
                    }
                    if (allIOsClone.Count > 0)
                    {
                        pinLine += " .0;";
                    }
                    else
                    {
                        pinLine += " ;";
                    }

                    if (diffCol > 0)
                    {
                        string preset0 = $"{"> Tset ",21}" + Regex.Match(pinLine, "> Tset (?<str>.*)").Groups["str"].ToString().Replace(" X", " 0");
                        sw.WriteLine(preset0);
                    }
                    sw.WriteLine(pinLine);
                }

                pinLine = $"{"> Tset ",21}";
                for (int iY = 0; iY < testCnt; iY++)
                {
                    if (iX == iY)
                    {
                        pinLine += " M";
                    }
                    else if (diffCol == iY)
                    {
                        pinLine += " X";
                    }
                    else
                    {
                        pinLine += " 0";
                    }
                }
                if (allIOsClone.Count > 0)
                {
                    pinLine += " .0;";
                }
                else
                {
                    pinLine += " ;";
                }

                sw.WriteLine(pinLine);
            }
        }

        private void WriteHeader(StreamWriter sw, List<string> allIOsClone, List<string> testIOs, string moduleName)
        {
            string header = "vm_vector\n" + moduleName + "($tset";
            #region print header
            foreach (string pin in testIOs)
            {
                header += ", " + pin;
                if (allIOsClone.Contains(pin))
                {
                    allIOsClone.Remove(pin);
                }
            }
            if (allIOsClone.Count > 0)
            {
                header += ", (";
                foreach (string pin in allIOsClone)
                {
                    header += pin + ",";
                }

                header += ",)";
                header = header.Replace(",,", "");
            }
            header += ")\n{\n";
            sw.WriteLine(header);
            #endregion
        }
    }
}
