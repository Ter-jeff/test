using System.Collections.Generic;
using System.IO;

using RfLib.Dvdc.Reader.CapturePostProcess;

namespace RfLib.Dvdc.Reader.DsscSetup
{
    public class DsscSetupSheet
    {
        #region Field
        private const string HeaderDsscSetup = "DsscSetup";
        private const string HeaderPattern = "Pattern";
        private const string HeaderDigSrcEqn = "DigSrcEqn";
        private const string HeaderDigSrcReg = "DigSrcReg";
        private const string HeaderDigSrcAssignment = "DigSrcAssignment";
        private const string HeaderDigSrcPin = "DigSrcPin";
        private const string HeaderDigSrcSampleSize = "DigSrcSampleSize";
        private const string HeaderDigCapPin = "DigCapPin";
        private const string HeaderDigCapSampleSize = "DigCapSampleSize";
        private const string HeaderCusStrDigCapData = "CusStrDigCapData";
        private const string HeaderPatModuleInfo = "PatModuleInfo";
        #endregion

        #region Properity
        public string SheetName { set; get; }

        public Dictionary<string, List<DsscSetupRow>> Setups = [];
        public Dictionary<string, List<DsscSetupRow>> BackupSetups = [];

        public List<DsscSetupSheetRow> RowList { set; get; }
        public List<DsscSetupSheetRow> InitRowList { set; get; }
        public List<PostProcessSheetRow> PostProcessRowList { set; get; }
        #endregion

        #region Constructor
        public DsscSetupSheet()
        {
            SheetName = "";
            RowList = [];
            InitRowList = [];
            PostProcessRowList = [];
        }

        public DsscSetupSheet(string sheet)
        {
            SheetName = sheet;
            RowList = [];
            InitRowList = [];
            PostProcessRowList = [];
        }
        #endregion

        #region MemberFunction

        public static Dictionary<string, DsscSetupSheetRow>? GetPatternDic()
        {
            return null;
        }

        public void Write(string path)
        {
            var sw = new StreamWriter(path);
            sw.WriteLine(string.Join("\t", GetHeaders()));
            foreach (KeyValuePair<string, List<DsscSetupRow>> setup in Setups)
            {
                foreach (DsscSetupRow dsscRow in setup.Value)
                {
                    sw.WriteLine(string.Join("\t", dsscRow.GetInfos()));
                }
            }
            if (BackupSetups.Count > 0)
            {
                sw.WriteLine();
                foreach (KeyValuePair<string, List<DsscSetupRow>> setup in BackupSetups)
                {
                    foreach (DsscSetupRow dsscRow in setup.Value)
                    {
                        sw.WriteLine(string.Join("\t", dsscRow.GetInfos()));
                    }
                }
            }

            sw.Close();
        }

        private static List<string> GetHeaders()
        {

            return [HeaderDsscSetup,
                HeaderPattern,
                HeaderDigSrcEqn,
                HeaderDigSrcReg,
                HeaderDigSrcAssignment,
                HeaderDigSrcPin,
                HeaderDigSrcSampleSize,
                HeaderDigCapPin,
                HeaderDigCapSampleSize,
                HeaderCusStrDigCapData,
                HeaderPatModuleInfo];
        }
        #endregion
    }
}
