using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using EfuseCheckCmdLib.IgxlLogLib.Base;

using LogLib.Utility;

namespace EfuseCheckCmdLib.IgxlLogLib
{
    public partial class DataLogFileInfo(string fullFileNamePath)
    {

        private const string SubKeyStrSheetandverinfo = "Sheet and Version Information";
        private const string SubKeyStrAutogenversion = "Autogen Version:";

        private const string SubKeyStrExec = "Site Failed tests/Executed tests";
        private const string SubKeyStrSbin = "Site    Sort     Bin";
        private const string SubKeyStrXy = "Site    X_Coord     Y_Coord";

        [GeneratedRegex(@"\s+|\t+")]
        private static partial Regex SpaceOrTabRegex();

        [GeneratedRegex("Prog Name")]
        private static partial Regex ProgNameRegex();

        [GeneratedRegex("(.xlsm)|(.xls)")]
        private static partial Regex XlsExtensionRegex();

        [GeneratedRegex(":")]
        private static partial Regex ColonRegex();

        [GeneratedRegex("Current job name")]
        private static partial Regex CurrentJobNameRegex();

        [GeneratedRegex("Current part type")]
        private static partial Regex CurrentPartTypeRegex();

        [GeneratedRegex(@"Active\s*EnableWords\s*:(?<enableWd>.*)")]
        private static partial Regex EnableWordsRegex();

        [GeneratedRegex(@"\s*Number\s*Site\s*Test\s*Name", RegexOptions.IgnoreCase)]
        private static partial Regex NumberSiteTestNameRegex();

        [GeneratedRegex(@"\s*Site\s*Failed\s*tests/Executed\s*tests", RegexOptions.IgnoreCase)]
        private static partial Regex SiteFailedTestsExecutedTestsRegex();

        [GeneratedRegex("\r\n")]
        private static partial Regex CrLfRegex();

        [GeneratedRegex("\n")]
        private static partial Regex LfRegex();

        [GeneratedRegex(@"^\s*\d+")]
        private static partial Regex LeadingDigitsRegex();

        [GeneratedRegex(@"\s+")]
        private static partial Regex WhitespaceRegex();

        [GeneratedRegex(SubKeyStrExec)]
        private static partial Regex SubKeyExecRegex();

        [GeneratedRegex(SubKeyStrSbin)]
        private static partial Regex SubKeySbinRegex();

        [GeneratedRegex(SubKeyStrXy)]
        private static partial Regex SubKeyXyRegex();

        [GeneratedRegex("-")]
        private static partial Regex HyphenRegex();

        private string FullFileName { get; } = fullFileNamePath;
        public string ProgramName { get; set; } = "";
        public string JobName { get; set; } = "";
        public string PartType { get; set; } = "";
        public List<int> CurrentDeviceNoList = [];
        private List<TouchDown> _tdPointers = [];

        public List<TouchDown> GeTouchDowns()
        {
            return _tdPointers;
        }

        public List<TouchDown> MGetDevicePointer()
        {
            _tdPointers = [];
            //Stopwatch stopWatch = new Stopwatch();
            long fileLength = new FileInfo(FullFileName).Length;
            long length = fileLength > 200000 ? 200000 : fileLength;
            //stopWatch.Start();

            bool firstTd = true;

            #region :: Get pointer of "Device#"
            using (FileStream stream = File.Open(FullFileName, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite))
            {
                using var mmf = MemoryMappedFile.CreateFromFile(stream, null, fileLength, MemoryMappedFileAccess.Read, HandleInheritability.Inheritable, false);
                long offset = 0;
                long segment = length - 1;
                while (offset + segment < fileLength + length)
                {
                    long segmentValue = offset + segment > fileLength ? fileLength - offset : segment;
                    using MemoryMappedViewStream viewStream = mmf.CreateViewStream(offset, segmentValue, MemoryMappedFileAccess.Read);
                    using (BinaryReader binReader = new BinaryReader(viewStream))
                    {
                        byte[] resultBytes = binReader.ReadBytes((int)segmentValue);
                        string result = Encoding.ASCII.GetString(resultBytes);
                        string keyStr = "Device#:";
                        int pointer = result.IndexOf(keyStr, StringComparison.Ordinal);
                        int pointerSheetAndVerInfo = result.IndexOf(SubKeyStrSheetandverinfo, StringComparison.Ordinal);
                        int pointerAutogenVersion = result.IndexOf(SubKeyStrAutogenversion, StringComparison.Ordinal);
                        int pointerSummaryStart = result.IndexOf(SubKeyStrExec, StringComparison.Ordinal);
                        int pointerSummaryEnd = result.IndexOf(SubKeyStrSbin, StringComparison.Ordinal);
                        if (pointer > 0 && offset == 0)
                        {
                            string lDatalogHeaderInfo = result[..pointer];
                            MParseDatalogHeaderInfo(lDatalogHeaderInfo);
                        }
                        if (pointerSheetAndVerInfo > 0 && pointerAutogenVersion > 0 && offset == 0)
                        {
                            string lDatalogSheetAndVerInfo = result[pointerSheetAndVerInfo..pointerAutogenVersion];
                            MParseDatalogSheetAndVerInfo(lDatalogSheetAndVerInfo);
                        }
                        while (pointer != -1)
                        {
                            var newTd = new TouchDown();
                            string device = Regex.Match(result[pointer..].Trim(), "(?<device>" + keyStr + ".*)").Groups["device"].ToString();
                            device = device.Trim();
                            newTd.RegionStartPtr = pointer + offset;
                            newTd.DeviceInfo = device;
                            if (firstTd)
                            {
                                newTd.FirstTd = firstTd;
                                firstTd = false;
                            }
                            _tdPointers.Add(newTd);
                            pointer = result.IndexOf(keyStr, pointer + 1, StringComparison.Ordinal);

                        }
                        while (_tdPointers.Count == 0 && pointerSummaryStart != -1 && pointerSummaryEnd != -1)
                        {
                            var newTd = new TouchDown();
                            string[] lines = result[pointerSummaryStart..pointerSummaryEnd].Replace("\r", "").Trim().Split('\n');
                            foreach (string line in lines)
                            {
                                if (string.IsNullOrWhiteSpace(line))
                                {
                                    continue;
                                }

                                if (line.StartsWith(SubKeyStrExec) || line.StartsWith("-----"))
                                {
                                    continue;
                                }

                                newTd.DeviceInfo = line.Trim().Split(' ')[0];
                                if (firstTd)
                                {
                                    newTd.FirstTd = firstTd;
                                    firstTd = false;
                                }
                                _tdPointers.Add(newTd);
                            }
                            pointerSummaryStart = result.IndexOf(SubKeyStrExec, pointerSummaryStart + 1, StringComparison.Ordinal);
                            pointerSummaryEnd = result.IndexOf(SubKeyStrSbin, pointerSummaryEnd + 1, StringComparison.Ordinal);
                        }
                    }
                    //prevent segment discontinuity issue
                    offset += segment - 10;
                }
            }
            for (int ipointer = 0; ipointer < _tdPointers.Count - 1; ipointer++)
            {
                _tdPointers[ipointer].SegmentSize = _tdPointers[ipointer + 1].RegionStartPtr - _tdPointers[ipointer].RegionStartPtr;
            }
            _tdPointers[^1].SegmentSize = fileLength - _tdPointers[^1].RegionStartPtr;
            #endregion

            using (var mmf = MemoryMappedFile.CreateFromFile(new FileStream(FullFileName, FileMode.Open), "myFile", fileLength, MemoryMappedFileAccess.ReadWrite, HandleInheritability.None, false))
            {
                #region :: Get line number of "Device#"
                long regionStartLine = 0;
                for (int ipointer = 0; ipointer < _tdPointers.Count; ipointer++)
                {
                    long offset;
                    long segment;
                    if (ipointer == 0 && _tdPointers[ipointer].RegionStartPtr > 0)
                    {
                        offset = 0;
                        segment = _tdPointers[ipointer].RegionStartPtr;
                        using MemoryMappedViewStream viewStream = mmf.CreateViewStream(offset, segment);
                        using var binReader = new BinaryReader(viewStream);
                        byte[] resultBytes = binReader.ReadBytes((int)segment);
                        regionStartLine = resultBytes.Where(a => a == '\n').ToList().Count;
                    }

                    _tdPointers[ipointer].RegionStartLine = regionStartLine + 1;
                    offset = _tdPointers[ipointer].RegionStartPtr;
                    segment = _tdPointers[ipointer].SegmentSize;
                    using (MemoryMappedViewStream viewStream = mmf.CreateViewStream(offset, segment))
                    {
                        using var binReader = new BinaryReader(viewStream);
                        int tmpsize = 1024 * 1024;
                        byte[] resultBytes = binReader.ReadBytes(tmpsize);
                        while (resultBytes.Length > 0)
                        {
                            int sectionLines = resultBytes.Where(a => a == '\n').ToList().Count;
                            _tdPointers[ipointer].RegionLines += sectionLines;
                            regionStartLine += sectionLines;
                            resultBytes = binReader.ReadBytes(tmpsize);
                        }
                    }
                }
                #endregion

                #region :: Read out the "Device#" information
                Parallel.ForEach(_tdPointers, td => MGetInfoOfEachTdParallel(mmf, td));
                #endregion
            }

            return _tdPointers;

        }

        private void MParseDatalogHeaderInfo(string sb) //Get 檔頭那些資訊
        {
            string[] tmpArray = sb.Trim().Split(["\r\n"], StringSplitOptions.RemoveEmptyEntries);

            //var keyLineNumber = -1;
            //var currLineNumber = 0;

            for (int i = 0; i < tmpArray.Length; i++)
            {
                //以space or \t Split
                string[] iArray = SpaceOrTabRegex().Split(tmpArray[i]);

                if (ProgNameRegex().IsMatch(tmpArray[i]))
                {
                    ProgramName = XlsExtensionRegex().Replace(iArray[^1], "");
                    if (ProgNameRegex().IsMatch(tmpArray[i]))
                    {
                        ProgramName = iArray[^1].Replace(".xls", "");
                    }
                }

            }
        }

        private void MParseDatalogSheetAndVerInfo(string sb) //Get 檔頭那些資訊
        {
            string[] tmpArray = sb.Trim().Split(["\r\n"], StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < tmpArray.Length; i++)
            {
                //以space or \t Split
                string[] iArray = ColonRegex().Split(tmpArray[i]);

                if (CurrentJobNameRegex().IsMatch(tmpArray[i]))
                {
                    JobName = iArray[^1];
                }

                if (CurrentPartTypeRegex().IsMatch(tmpArray[i]))
                {
                    PartType = iArray[^1];
                }
            }
        }

        private void MGetInfoOfEachTdParallel(MemoryMappedFile memoryMappedFile, TouchDown touchDown)
        {
            string? sb;
            long offset = touchDown.RegionStartPtr;
            long segment = touchDown.SegmentSize;

            using (MemoryMappedViewStream viewStream = memoryMappedFile.CreateViewStream(offset, segment))
            {
                using var binReader = new BinaryReader(viewStream);
                // in TD content, offset is no require any more.

                long tdSegment = 0, tdOffset = 0;
                // 500000
                const long regional = 500000;
                tdSegment = segment > regional ? regional : segment;
                // only take few lines to do search

                #region // Get the Func and Meas Column Arrangement
                sb = Encoding.ASCII.GetString(binReader.ReadBytes((int)tdSegment));
                string subKeyStr = "Device#:";
                string device = Regex.Match(sb, "(?<device>" + subKeyStr + ".*)").Groups["device"].ToString().Trim();
                int readSegCnt = 1;
                #endregion

                #region Get Enable Words

                string enableWd = EnableWordsRegex().Match(sb).Groups["enableWd"].ToString().Trim();
                bool searchFlag = string.IsNullOrEmpty(enableWd);
                while (searchFlag)
                {
                    if (segment > regional)
                    {
                        tdSegment = segment > regional ? regional : segment;
                        sb = Encoding.ASCII.GetString(binReader.ReadBytes((int)tdSegment));
                    }
                    enableWd = EnableWordsRegex().Match(sb).Groups["enableWd"].ToString().Trim();

                    if (NumberSiteTestNameRegex().IsMatch(sb))
                    {
                        searchFlag = false;
                    }

                    if (SiteFailedTestsExecutedTestsRegex().IsMatch(sb))
                    {
                        searchFlag = false;
                    }
                }
                touchDown.EnableWords = enableWd.Replace("'", "");

                #endregion

                #region // Get the regional pointers
                subKeyStr = SubKeyStrExec;
                tdOffset = segment - regional > 0 ? segment - regional : 0;

                binReader.BaseStream.Position = tdOffset;
                // only take few lines to do search
                sb = Encoding.ASCII.GetString(binReader.ReadBytes((int)tdSegment));
                int subPointer = -1;
                readSegCnt = 1;
                while (subPointer == -1)
                {
                    subPointer = sb.IndexOf(subKeyStr, 0, StringComparison.Ordinal);
                    if (subPointer == -1)
                    {
                        readSegCnt++;
                        tdOffset = tdOffset - (regional * (readSegCnt - 1)) < 0 ? 0 : tdOffset - (regional * (readSegCnt - 1));
                        binReader.BaseStream.Position = tdOffset;
                        // only take few lines to do search
                        sb = Encoding.ASCII.GetString(binReader.ReadBytes((int)tdSegment));
                        if (tdOffset == 0)
                        {
                            ErrorMessageBox.Show("Error, can't find pair string \"Site Failed tests/Executed tests\" in the log", "Format Error");
                            return;
                        }
                    }
                }

                if (string.IsNullOrEmpty(device))
                {
                    touchDown.CurrDeviceNum.Add(int.Parse(touchDown.DeviceInfo));
                }
                else
                {
                    touchDown.CurrDeviceNum.AddRange(ParseDeviceNumber(device));
                }

                touchDown.ExecuteTestPtr = subPointer + tdOffset;
                string subSb = sb[subPointer..(int)tdSegment];
                int subPointerExec = subSb.IndexOf(SubKeyStrSbin, 0, StringComparison.Ordinal);
                touchDown.ExecuteTestSize = subPointerExec;

                subPointer = sb.IndexOf(SubKeyStrSbin, 0, StringComparison.Ordinal);
                touchDown.SortBinPtr = subPointer + tdOffset;
                subSb = sb[subPointer..(int)tdSegment];
                int subPointerSBin = subSb.IndexOf(SubKeyStrXy, 0, StringComparison.Ordinal);
                touchDown.SortBinSize = subPointerSBin;

                subKeyStr = SubKeyStrXy;
                subPointer = sb.IndexOf(subKeyStr, 0, StringComparison.Ordinal);
                touchDown.DieXyPtr = subPointer + tdOffset;
                subSb = sb[subPointer..(int)tdSegment];
                //var subPointerXY = subSb.IndexOf("---", 0, StringComparison.Ordinal);

                // 嘗試找 ===，如果找不到就找 Device#，都沒有就用 sb 結尾
                int subPointerXy = subSb.IndexOf("===", 0, StringComparison.Ordinal);
                if (subPointerXy < 0)
                {
                    // 找 Device#:
                    int deviceIndex = subSb.IndexOf("Device#:", StringComparison.OrdinalIgnoreCase);
                    if (deviceIndex >= 0)
                    {
                        subPointerXy = deviceIndex;
                    }
                    else
                    {
                        // fallback 到 EOF
                        subPointerXy = subSb.Length;
                    }
                }

                touchDown.DieXySize = subPointerXy;

                binReader.BaseStream.Position = touchDown.ExecuteTestPtr;
                subSb = Encoding.ASCII.GetString(binReader.ReadBytes((int)touchDown.ExecuteTestSize));
                MParseSummary(touchDown, subSb);

                binReader.BaseStream.Position = touchDown.SortBinPtr;
                subSb = Encoding.ASCII.GetString(binReader.ReadBytes((int)touchDown.SortBinSize));
                MParseSummary(touchDown, subSb);

                binReader.BaseStream.Position = touchDown.DieXyPtr;
                string text = ReadNextLines(viewStream, 3);
                //subSb = Encoding.ASCII.GetString(binReader.ReadBytes((int)newTd.DieXySize));
                MParseSummary(touchDown, text);

                if (touchDown.DieXyPtr >= 0 &&
                    touchDown.DieXySize > 0 &&
                    binReader.BaseStream.Length >= touchDown.DieXyPtr + touchDown.DieXySize)
                {
                    binReader.BaseStream.Position = touchDown.DieXyPtr;
                    subSb = Encoding.ASCII.GetString(binReader.ReadBytes((int)touchDown.DieXySize));
                }
                else
                {
                    // 讀不到就給N/A的，後面 MParseSummary 可處理為 X=0, Y=0
                    subSb = "N/A";
                }

                MParseSummary(touchDown, subSb);
                int idx = 0;
                if (!MCheckSiteCount(touchDown))
                {
                    return;
                }

                foreach (DieInfo die in touchDown.DieTestSum)
                {
                    die.DeviceNum = touchDown.CurrDeviceNum[idx];
                    idx++;
                }
                //lDevices.Add(device);
                #endregion

                foreach (int dNum in touchDown.CurrDeviceNum)
                {
                    CurrentDeviceNoList.Add(dNum);
                }

                sb = null;
            }
            GC.Collect();
        }

        public static string ReadNextLines(Stream stream, int numLines)
        {
            StringBuilder nextLines = new StringBuilder();
            using (StreamReader reader = new StreamReader(stream, Encoding.ASCII, false, 1024, true))
            {
                for (int i = 0; i < numLines; i++)
                {
                    string? line = reader.ReadLine();
                    if (string.IsNullOrEmpty(line))
                    {
                        // Stop if EOF
                        break;
                    }

                    nextLines.AppendLine(line);
                }
            }
            return nextLines.ToString();
        }

        private static bool MCheckSiteCount(TouchDown touchDown)
        {
            if (touchDown.CurrActiveSiteNum.Count != touchDown.CurrDeviceNum.Count)
            {
                ErrorMessageBox.Show($"Site Number {touchDown.CurrDeviceNum.Count} != Device Number {touchDown.CurrActiveSiteNum.Count}");
                return false;
            }
            return true;
        }

        private static void MParseSummary(TouchDown touchDown, string line)
        {
            string[] tmp = CrLfRegex().Split(line.Trim());
            if (tmp.Length == 1)
            {
                tmp = LfRegex().Split(line.Trim());
            }

            int site = 0;
            string rt1 = "NA", rt2 = "NA";
            bool start = false;
            foreach (string item in tmp)
            {
                if (LeadingDigitsRegex().IsMatch(item))
                {
                    start = true;
                    string[] dieTest = WhitespaceRegex().Split(item.Trim());
                    site = Convert.ToInt32(dieTest[0]);
                    rt1 = dieTest[1];
                    rt2 = dieTest[2];
                    DieInfo? tdExisted = null;
                    if (touchDown.DieTestSum.Count > 0)
                    {
                        var lfound = touchDown.DieTestSum.Where(w => w.Site == site).Select(g => g).ToList();
                        if (lfound.Count > 0)
                        {
                            tdExisted = lfound.First();
                        }
                    }
                    // if tdExisted == null  
                    DieInfo dinfo = tdExisted ?? new DieInfo();
                    dinfo.Site = site;
                    if (!touchDown.CurrActiveSiteNum.Contains(site))
                    {
                        touchDown.CurrActiveSiteNum.Add(site);
                    }

                    if (SubKeyExecRegex().IsMatch(line))
                    {
                        dinfo.ExecFailTests = rt1;
                        dinfo.ExecTests = rt2;
                    }
                    else if (SubKeySbinRegex().IsMatch(line))
                    {
                        dinfo.Sort = rt1;
                        dinfo.Bin = rt2;
                    }
                    else if (SubKeyXyRegex().IsMatch(line))
                    {
                        dinfo.XCoord = rt1;
                        dinfo.YCoord = rt2;
                    }
                    if (tdExisted == null)
                    {
                        touchDown.DieTestSum.Add(dinfo);
                    }
                }
                else if (start)
                {
                    start = false;
                    break;
                }
            }
        }

        private static List<int> ParseDeviceNumber(string line) //Device#: 1-4
        {
            var devList = new List<int>();
            string[] devices = line.Replace("Device#:", "").Trim().Split(',');

            foreach (string dev in devices)
            {
                if (HyphenRegex().IsMatch(dev))
                {
                    string[] tmpAry = dev.Split('-');
                    for (short t = Convert.ToInt16(tmpAry[0]); t <= Convert.ToInt16(tmpAry[1]); t++)
                    {
                        //沒有檢查能不能轉INT有點風險
                        devList.Add(t);
                    }
                }
                else
                {
                    //沒有檢查能不能轉INT有點風險
                    devList.Add(Convert.ToInt16(dev));
                }
            }
            return devList;
        }

    }
}
