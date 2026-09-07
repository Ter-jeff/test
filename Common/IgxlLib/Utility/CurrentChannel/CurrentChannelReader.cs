using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

namespace IgxlLib.Utility.CurrentChannel
{
    public class CurrentChannelReader
    {
        public static CurrentChannelSheet? ReadFile(string path)
        {
            var currentChannelSheet = new CurrentChannelSheet();
            Dictionary<string, List<CurrentChannelMapRow>> currentChannel;
            try
            {
                if (!File.Exists(path))
                {
                    return null;
                }

                bool startFlag = false;
                string instrumentName = "";
                currentChannel = [];
                var rows = new List<CurrentChannelMapRow>();
                using var sr = new StreamReader(path);
                while (!sr.EndOfStream)
                {
                    string? line = sr.ReadLine();

                    if (!startFlag)
                    {
                        if (line != null && line.StartsWithIgnoreCase("TesterChannel	DibChannel	SignalName"))
                        {
                            startFlag = true;
                            //TesterChannel	DibChannel	SignalName
                            sr.ReadLine();
                            //empty line
                            line = sr.ReadLine();
                        }
                    }

                    bool nextFlag = false;
                    if (startFlag)
                    {
                        if (line != null)
                        {
                            string[] lineSpt = line.Split('\t');
                            if (lineSpt.Length == 1 && !string.IsNullOrEmpty(lineSpt[0]))
                            {
                                instrumentName = lineSpt[0];
                            }

                            var row = new CurrentChannelMapRow();
                            if (lineSpt.Length == 5)
                            {
                                row.TesterChannel = lineSpt[0];
                                row.DibChannel = lineSpt[2];
                                row.SignalName = lineSpt[4];
                                rows.Add(row);
                            }
                        }

                        if (string.IsNullOrEmpty(line))
                        {
                            nextFlag = true;
                        }
                    }

                    if (nextFlag)
                    {
                        if (!string.IsNullOrEmpty(instrumentName) && !currentChannel.ContainsKey(instrumentName))
                        {
                            var newRow = rows.Select(x => x.Copy()).ToList();
                            currentChannel.Add(instrumentName, newRow);
                            rows.Clear();
                        }
                    }
                }
                if (!string.IsNullOrEmpty(instrumentName) && !currentChannel.ContainsKey(instrumentName))
                {
                    var channelMapRows = rows.Select(x => x.Copy()).ToList();
                    currentChannel.Add(instrumentName, channelMapRows);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Read current channel file failed for {path}: {ex.Message}", ex);
            }

            currentChannelSheet.Dic = currentChannel;
            return currentChannelSheet;
        }
    }

    public class CurrentChannelSheet
    {
        public Dictionary<string, List<CurrentChannelMapRow>>? Dic { get; internal set; }

        public Dictionary<string, string> GetPogoMapping(string group)
        {
            // 1. Check if the group exists in the source dictionary
            if (string.IsNullOrEmpty(group) || Dic == null || !Dic.TryGetValue(group, out List<CurrentChannelMapRow>? currentChannelMap))
            {
                return new Dictionary<string, string>(StringExtensions.IgnoreCase);
            }

            var result = new Dictionary<string, string>(StringExtensions.IgnoreCase);

            if (currentChannelMap == null)
            {
                return result;
            }

            foreach (CurrentChannelMapRow row in currentChannelMap)
            {
                // 2. Validate DibChannel and SignalName for nulls and the presence of a '.'
                if (string.IsNullOrEmpty(row.DibChannel) || !row.DibChannel.Contains('.') ||
                    string.IsNullOrEmpty(row.SignalName) || !row.SignalName.Contains('.'))
                {
                    // Skip malformed rows
                    continue;
                }

                string[] dibParts = row.DibChannel.Split('.');
                string[] signalParts = row.SignalName.Split('.');

                // 3. Ensure the array has at least 2 parts before accessing index [1]
                if (dibParts.Length > 1 && signalParts.Length > 1)
                {
                    string key = dibParts[1];
                    string value = signalParts[1];

                    // 4. Prevent ArgumentException (duplicate keys)
                    result.TryAdd(key, value);
                }
            }

            return result;
        }
    }
}
