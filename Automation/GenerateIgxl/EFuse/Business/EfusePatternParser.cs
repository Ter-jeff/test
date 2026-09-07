using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfusePatternParser
    {
        public int SendCount;
        public int StoreCount;
        public string ReadWritePin;

        private readonly string _patternPath;
        private readonly bool _isRead;
        private readonly bool _isWrite;

        public EfusePatternParser(string patternPath, bool isRead, bool isWrite)
        {
            _patternPath = patternPath;
            _isRead = isRead;
            _isWrite = isWrite;
        }

        public void WorkFlow()
        {
            SendCount = 0;
            StoreCount = 0;
            ReadWritePin = "";
            ReadWritePin = GetDigSrcCapPin();
            if (_isWrite && !string.IsNullOrEmpty(ReadWritePin))
            {
                SendCount = GetWritePatSendCount();
            }
            else if (_isRead && !string.IsNullOrEmpty(ReadWritePin))
            {
                StoreCount = GetReadPatStoreCount();
            }
        }

        private string GetDigSrcCapPin()
        {
            using (FileStream fs = File.OpenRead(_patternPath))
            {
                using (var zip = new GZipStream(fs, CompressionMode.Decompress))
                {
                    using (var sr = new StreamReader(zip))
                    {
                        var instruments = new Regex(@"instruments\s*=\s*\{", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                        var end = new Regex(@"[\}]", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                        bool isInstrument = false;
                        var instrumentsList = new List<string>();
                        while (!sr.EndOfStream)
                        {
                            string line = sr.ReadLine()?.Trim();
                            if (string.IsNullOrEmpty(line) || line.StartsWith("//"))
                            {
                                continue;
                            }

                            if (instruments.IsMatch(line))
                            {
                                isInstrument = true;
                            }

                            if (isInstrument)
                            {
                                instrumentsList.Add(line);
                            }

                            if (end.IsMatch(line))
                            {
                                break;
                            }
                        }

                        string instrumentStr = string.Join("\n", instrumentsList);
                        if (_isRead)
                        {
                            Match match = Regex.Match(instrumentStr, @"instruments\s*=\s*\{[\S|\s]*(([^\/\/]+\((?<pin>\S*)\):DigCap[^;]+))",
                                RegexOptions.IgnoreCase);
                            if (match.Success)
                            {
                                return match.Groups["pin"].ToString().ToUpper();
                            }
                        }

                        if (_isWrite)
                        {
                            Match match = Regex.Match(instrumentStr, @"instruments\s*=\s*\{[\S|\s]*(([^\/\/]+\((?<pin>\S*)\):DigSrc[^;]+))",
                                RegexOptions.IgnoreCase);
                            if (match.Success)
                            {
                                return match.Groups["pin"].ToString().ToUpper();
                            }
                        }
                    }
                }
            }

            return "";
        }

        private int GetWritePatSendCount()
        {
            int sendCount = 0;
            using (FileStream fs = File.OpenRead(_patternPath))
            {
                using (var zip = new GZipStream(fs, CompressionMode.Decompress))
                {
                    using (var sr = new StreamReader(zip))
                    {
                        while (!sr.EndOfStream)
                        {
                            string line = sr.ReadLine()?.Trim();

                            if (string.IsNullOrEmpty(line) || line.StartsWith("//"))
                            {
                                continue;
                            }

                            if (line.StartsWith("global subr", StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }

                            if (Regex.IsMatch(line, ReadWritePin + @"\s*\)\s*:\s*DigSrc\s*=\s*Send\s*\)", RegexOptions.IgnoreCase))
                            {
                                sendCount++;
                            }

                        }
                    }
                }
            }
            return sendCount;
        }

        private int GetReadPatStoreCount()
        {
            int storeCount = 0;
            using (FileStream fs = File.OpenRead(_patternPath))
            {
                using (var zip = new GZipStream(fs, CompressionMode.Decompress))
                {
                    using (var sr = new StreamReader(zip))
                    {
                        while (!sr.EndOfStream)
                        {
                            string line = sr.ReadLine()?.Trim();
                            if (string.IsNullOrEmpty(line) || line.StartsWith("//"))
                            {
                                continue;
                            }

                            if (line.StartsWith("global subr", StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }

                            if (Regex.IsMatch(line, ReadWritePin + @"\s*\)\s*:\s*DigCap\s*=\s*Store\s*\)", RegexOptions.IgnoreCase))
                            {
                                storeCount++;
                            }

                        }
                    }
                }
            }
            return storeCount;
        }
    }
}
