using System;
using System.Collections.Generic;
using System.Linq;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility
{
    public class PatternResult
    {
        public string FilePath { get; set; }
        public string Pattern { get; set; }
        public DateTime LastWriteTime { get; set; }
        public string TimeSet { get; set; }//other
        public string UseNoUse { get; set; }//other
        public bool HasPat { get; set; }//PAT
        public bool HasAtp { get; set; }//ATP
        public string VmVectorName { get; set; }//PAT
        public string OpcodeMode { get; set; }//PAT
        public string SubFunc { get; set; }//PAT

        public double ExpectedTime { get; set; }//other
        public int PatternVectorCount;//PAT
        public string TsetSummary { get; set; }//ATP
        public int TsetTotalCount { get; set; }//ATP
        public double ComtCount { get; set; }//ATP

        public int NumOfSvm { get; set; }//PAT
        public int NumOfLvm { get; set; }//PAT
        public int LicenseCount { get; set; }//PAT

        public int PinCount { get; set; }//PAT
        public string Pins { get; set; }//PAT

        public int MeasVCount { get; set; }//ATP
        public int MeasVdiffCount { get; set; }//ATP
        public int MeasVdiff1Count { get; set; }//ATP
        public int MeasVdiff2Count { get; set; }//ATP
        public int MeasVocmCount { get; set; }//ATP
        public int MeasICount { get; set; }//ATP
        public int MeasIDiffCount { get; set; }//ATP
        public int MeasIDiff1Count { get; set; }//ATP
        public int MeasIDiff2Count { get; set; }//ATP
        public int MeasI2Count { get; set; }//ATP
        public int MeasR1Count { get; set; }//ATP
        public int MeasR2Count { get; set; }//ATP
        public int MeasFCount { get; set; }//ATP
        public int MeasFDiffCount { get; set; }//ATP
        public int MeasFDiff1Count { get; set; }//ATP
        public int MeasFDiff2Count { get; set; }//ATP
        public int MeasCCount { get; set; }//ATP
        public int MeasECount { get; set; }//ATP

        public string ProcessTime { get; set; }

        public List<string> PinList = new List<string>();//PAT
        //public List<ScanMap> ScanMapList = new List<ScanMap>();//PAT
        //public List<Unresolved> UnresolvedList = new List<Unresolved>();//PAT
        public List<string> ModuleNameList = new List<string>();//PAT
        public Dictionary<string, string> ScanInList = new Dictionary<string, string>();//PAT
        public Dictionary<string, int> TsetCountSummary = new Dictionary<string, int>();//ATP

        public string GetPatternType()
        {
            if (Pattern == null)
            {
                return "";
            }

            if (Pattern.Contains("_"))
            {
                return Pattern.Split('_').First().ToUpper();
            }

            return Pattern.ToLower();
        }

        public AtpResult CovertAtpResult()
        {
            var atpResult = new AtpResult
            {
                FilePath = FilePath, Pattern = Pattern, LastWriteTime = LastWriteTime, HasAtp = HasAtp,
                TsetSummary = TsetSummary,
                TsetTotalCount = TsetTotalCount,
                ComtCount = ComtCount,
                MeasVCount = MeasVCount,
                MeasVdiffCount = MeasVdiffCount,
                MeasVdiff1Count = MeasVdiff1Count,
                MeasVdiff2Count = MeasVdiff2Count,
                MeasVocmCount = MeasVocmCount,
                MeasICount = MeasICount,
                MeasIDiffCount = MeasIDiffCount,
                MeasIDiff1Count = MeasIDiff1Count,
                MeasIDiff2Count = MeasIDiff2Count,
                MeasI2Count = MeasI2Count,
                MeasR1Count = MeasR1Count,
                MeasR2Count = MeasR2Count,
                MeasFCount = MeasFCount,
                MeasFDiffCount = MeasFDiffCount,
                MeasFDiff1Count = MeasFDiff1Count,
                MeasFDiff2Count = MeasFDiff2Count,
                MeasCCount = MeasCCount,
                MeasECount = MeasECount,
                TsetCountSummary = TsetCountSummary
            };
            return atpResult;
        }

        public PatResult CovertPatResult()
        {
            var patResult = new PatResult
            {
                FilePath = FilePath, Pattern = Pattern, LastWriteTime = LastWriteTime, HasPat = HasPat,
                VmVectorName = VmVectorName,
                OpcodeMode = OpcodeMode,
                SubFunc = SubFunc,
                PatternVectorCount = PatternVectorCount,
                NumOfSvm = NumOfSvm,
                NumOfLvm = NumOfLvm,
                LicenseCount = LicenseCount,
                PinCount = PinCount,
                Pins = Pins,
                PinList = PinList,
                //patResult.ScanMapList = ScanMapList;
                //patResult.UnresolvedList = UnresolvedList;
                ModuleNameList = ModuleNameList,
                ScanInList = ScanInList
            };
            return patResult;
        }

        public void GetValueBypAtpResult(AtpResult atpResult)
        {
            FilePath = atpResult.FilePath;
            Pattern = atpResult.Pattern;
            LastWriteTime = atpResult.LastWriteTime;
            HasAtp = atpResult.HasAtp;
            TsetSummary = atpResult.TsetSummary;
            TsetTotalCount = atpResult.TsetTotalCount;
            ComtCount = atpResult.ComtCount;
            MeasVCount = atpResult.MeasVCount;
            MeasVdiffCount = atpResult.MeasVdiffCount;
            MeasVdiff1Count = atpResult.MeasVdiff1Count;
            MeasVdiff2Count = atpResult.MeasVdiff2Count;
            MeasVocmCount = atpResult.MeasVocmCount;
            MeasICount = atpResult.MeasICount;
            MeasIDiffCount = atpResult.MeasIDiffCount;
            MeasIDiff1Count = atpResult.MeasIDiff1Count;
            MeasIDiff2Count = atpResult.MeasIDiff2Count;
            MeasI2Count = atpResult.MeasI2Count;
            MeasR1Count = atpResult.MeasR1Count;
            MeasR2Count = atpResult.MeasR2Count;
            MeasFCount = atpResult.MeasFCount;
            MeasFDiffCount = atpResult.MeasFDiffCount;
            MeasFDiff1Count = atpResult.MeasFDiff1Count;
            MeasFDiff2Count = atpResult.MeasFDiff2Count;
            MeasCCount = atpResult.MeasCCount;
            MeasECount = atpResult.MeasECount;
            TsetCountSummary = atpResult.TsetCountSummary;
        }

        public void GetValueBypPatResult(PatResult patResult)
        {
            FilePath = patResult.FilePath;
            Pattern = patResult.Pattern;
            LastWriteTime = patResult.LastWriteTime;
            HasPat = patResult.HasPat;
            VmVectorName = patResult.VmVectorName;
            OpcodeMode = patResult.OpcodeMode;
            SubFunc = patResult.SubFunc;
            PatternVectorCount = patResult.PatternVectorCount;
            NumOfSvm = patResult.NumOfSvm;
            NumOfLvm = patResult.NumOfLvm;
            LicenseCount = patResult.LicenseCount;
            PinCount = patResult.PinCount;
            Pins = patResult.Pins;
            PinList = patResult.PinList;
            //ScanMapList = patResult.ScanMapList;
            //UnresolvedList = patResult.UnresolvedList;
            ModuleNameList = patResult.ModuleNameList;
            ScanInList = patResult.ScanInList;
        }
    }

    public class PatResult
    {
        public string FilePath { get; set; }
        public string Pattern { get; set; }
        public DateTime LastWriteTime { get; set; }
        public bool HasPat { get; set; }//PAT
        public string VmVectorName { get; set; }//PAT
        public string OpcodeMode { get; set; }//PAT
        public string SubFunc { get; set; }//PAT
        public int PatternVectorCount;//PAT
        public int NumOfSvm { get; set; }//PAT
        public int NumOfLvm { get; set; }//PAT
        public int LicenseCount { get; set; }//PAT
        public int PinCount { get; set; }//PAT
        public string Pins { get; set; }//PAT
        public List<string> PinList = new List<string>();//PAT
        //public List<ScanMap> ScanMapList = new List<ScanMap>();//PAT
        //public List<Unresolved> UnresolvedList = new List<Unresolved>();//PAT
        public List<string> ModuleNameList = new List<string>();//PAT
        public Dictionary<string, string> ScanInList = new Dictionary<string, string>();//PAT
    }

    public class AtpResult
    {
        public string FilePath { get; set; }
        public string Pattern { get; set; }
        public DateTime LastWriteTime { get; set; }
        public bool HasAtp { get; set; }//ATP
        public string TsetSummary { get; set; }//ATP
        public int TsetTotalCount { get; set; }//ATP
        public double ComtCount { get; set; }//ATP
        public int MeasVCount { get; set; }//ATP
        public int MeasVdiffCount { get; set; }//ATP
        public int MeasVdiff1Count { get; set; }//ATP
        public int MeasVdiff2Count { get; set; }//ATP
        public int MeasVocmCount { get; set; }//ATP
        public int MeasICount { get; set; }//ATP
        public int MeasIDiffCount { get; set; }//ATP
        public int MeasIDiff1Count { get; set; }//ATP
        public int MeasIDiff2Count { get; set; }//ATP
        public int MeasI2Count { get; set; }//ATP
        public int MeasR1Count { get; set; }//ATP
        public int MeasR2Count { get; set; }//ATP
        public int MeasFCount { get; set; }//ATP
        public int MeasFDiffCount { get; set; }//ATP
        public int MeasFDiff1Count { get; set; }//ATP
        public int MeasFDiff2Count { get; set; }//ATP
        public int MeasCCount { get; set; }//ATP
        public int MeasECount { get; set; }//ATP
        public Dictionary<string, int> TsetCountSummary = new Dictionary<string, int>();//ATP
    }
}
