using System;
using System.Collections.Generic;
using System.IO;

using LogLib.Utility;

namespace Automation.Static
{
    public static class FolderStructure
    {

        #region Const Field
        private const string StrIgLink = "IGLink";
        private const string StrLogs = "Logs";
        private const string StrCommon = "Common";
        private const string StrModule = "Module";
        private const string StrXmlFiles = "xml_Files";
        private const string StrAcSpec = "AC_Spec";
        private const string StrBinTbl = "BinTable";
        private const string StrChannelMap = "Channel_Map";
        private const string StrCommonSheets = "Common_Sheets";
        private const string StrCommonVbt = "Common_VBT";
        private const string StrDcSpec = "DC_Spec";
        private const string StrDevChar = "DevChar";
        private const string StrGlbSpec = "Global_Spec";
        private const string StrLevel = "Levels";
        private const string StrPinMap = "PinMap";
        private const string StrPorts = "Ports";
        private const string StrTimings = "Timings";
        private const string StrJob = "Jobs";
        private const string StrPatSetsAll = "PatSetsAll";
        private const string StrEfuse = "eFuse";
        private const string StrScan = "Scan";
        private const string StrSocScan = "socScan";
        private const string StrSocMbist = "socMbist";
        private const string StrCpuScan = "cpuScan";
        private const string StrMbist = "Mbist";
        private const string StrCpuMbist = "cpuMbist";
        private const string StrGpuScan = "gfxScan";
        private const string StrGpuMbist = "gfxMbist";
        private const string StrHardIp = "HardIP";
        private const string StrLib = "Library";
        private const string StrCodingLib = "LibraryCoding";
        private const string StrDirVbtGenTool = "VBTGenTool";
        private const string StrBinCut = "BinCut";
        private const string StrNonBinCut = "NonBinCut";
        private const string StrEvs = "Evs";
        private const string StrCpm = "Cpm";
        private const string StrSpi = "Spi";
        private const string StrConti = "DC_Conti";
        private const string StrSpiRom = "SPI_ROM";
        private const string StrRtos = "Rtos";
        private const string StrMain = "Main";
        private const string StrAnalog = "Analog";
        private const string StrReference = "Reference";
        private const string StrSubProgram = "Subprogram";
        private const string StrT0Tx = "T0Tx";
        private const string StrVreValid = "VRE_ValidationList";
        #endregion

        //Folder1
        private const string SubDirIgLink = StrIgLink;
        //Folder2
        private static readonly string _subDirXmlFiles = Path.Combine(SubDirIgLink, StrXmlFiles);
        private static readonly string _subDirLogs = Path.Combine(SubDirIgLink, StrLogs);
        private static readonly string _subDirSubProgram = Path.Combine(SubDirIgLink, StrSubProgram);
        private static readonly string _subDirT0Tx = Path.Combine(SubDirIgLink, StrT0Tx);
        private static readonly string _subDirVreValid = Path.Combine(SubDirIgLink, StrVreValid);
        //Folder3
        private static readonly string _subDirCommon = Path.Combine(SubDirIgLink, StrCommon);
        //Folder4
        private static readonly string _subDirAcSpec = Path.Combine(_subDirCommon, StrAcSpec);
        private static readonly string _subDirBinTable = Path.Combine(_subDirCommon, StrBinTbl);
        private static readonly string _subDirChannelMap = Path.Combine(_subDirCommon, StrChannelMap);
        private static readonly string _subDirCommonSheets = Path.Combine(_subDirCommon, StrCommonSheets);
        private static readonly string _subDirCommonVbt = Path.Combine(_subDirCommon, StrCommonVbt);
        private static readonly string _subDirDcSpec = Path.Combine(_subDirCommon, StrDcSpec);
        private static readonly string _subDirDevChar = Path.Combine(_subDirCommon, StrDevChar);
        private static readonly string _subDirGlbSpec = Path.Combine(_subDirCommon, StrGlbSpec);
        private static readonly string _subDirLevel = Path.Combine(_subDirCommon, StrLevel);
        private static readonly string _subDirPinMap = Path.Combine(_subDirCommon, StrPinMap);
        private static readonly string _subDirPorts = Path.Combine(_subDirCommon, StrPorts);
        private static readonly string _subDirTimings = Path.Combine(_subDirCommon, StrTimings);
        private static readonly string _subDirJob = Path.Combine(_subDirCommon, StrJob);
        private static readonly string _subDirPatSetsAll = Path.Combine(_subDirCommon, StrPatSetsAll);
        private static readonly string _subDirReference = Path.Combine(_subDirCommon, StrReference);

        //Folder3
        private static readonly string _subDirModule = Path.Combine(SubDirIgLink, StrModule);
        //Folder4
        private static readonly string _subDirEfuse = Path.Combine(_subDirModule, StrEfuse);
        private static readonly string _subDirScan = Path.Combine(_subDirModule, StrScan);
        private static readonly string _subDirSocScan = Path.Combine(_subDirModule, StrSocScan);
        private static readonly string _subDirSocMbist = Path.Combine(_subDirModule, StrSocMbist);
        private static readonly string _subDirCpuScan = Path.Combine(_subDirModule, StrCpuScan);
        private static readonly string _subDirMbist = Path.Combine(_subDirModule, StrMbist);
        private static readonly string _subDirCpuMbist = Path.Combine(_subDirModule, StrCpuMbist);
        private static readonly string _subDirGpuScan = Path.Combine(_subDirModule, StrGpuScan);
        private static readonly string _subDirGpuMbist = Path.Combine(_subDirModule, StrGpuMbist);
        private static readonly string _subDirHardIp = Path.Combine(_subDirModule, StrHardIp);

        private static readonly string _subDirLib = Path.Combine(_subDirModule, StrLib);
        private static readonly string _subDirCodingLib = Path.Combine(_subDirModule, StrCodingLib);
        private static readonly string _subDirVbtGenTool = Path.Combine(_subDirModule, StrDirVbtGenTool);
        private static readonly string _subDirBinCut = Path.Combine(_subDirModule, StrBinCut);
        private static readonly string _subDirNonBinCut = Path.Combine(_subDirModule, StrNonBinCut);
        private static readonly string _subDirEvs = Path.Combine(_subDirModule, StrEvs);
        private static readonly string _subDirCpm = Path.Combine(_subDirModule, StrCpm);
        private static readonly string _subDirSpi = Path.Combine(_subDirModule, StrSpi);
        private static readonly string _subDirConti = Path.Combine(_subDirModule, StrConti);
        private static readonly string _subDirSpiRom = Path.Combine(_subDirModule, StrSpiRom);
        private static readonly string _subDirRtos = Path.Combine(_subDirModule, StrRtos);
        private static readonly string _subDirMain = Path.Combine(_subDirModule, StrMain);
        private static readonly string _subDirAnalog = Path.Combine(_subDirModule, StrAnalog);
        //Folder5
        private static List<string> _dirList;
        private static string _tarDir;

        #region Property
        public static string TarDir
        {
            get { return string.IsNullOrEmpty(_tarDir) && !string.IsNullOrEmpty(LocalSpecs.TarFolder) ? LocalSpecs.TarFolder : _tarDir; }
            set { _tarDir = value; }
        }

        public static string DirIgLink
        {
            get { return Contact(SubDirIgLink); }
        }

        public static string DirLogs
        {
            get { return Contact(_subDirLogs); }
        }

        public static string DirXmlFiles
        {
            get { return Contact(_subDirXmlFiles); }
        }

        public static string DirCommon
        {
            get { return Contact(_subDirCommon); }
        }

        public static string DirAcSpec
        {
            get { return Contact(_subDirAcSpec); }
        }

        public static string DirBinTable
        {
            get { return Contact(_subDirBinTable); }
        }

        public static string DirChannelMap
        {
            get { return Contact(_subDirChannelMap); }
        }

        public static string DirCommonSheets
        {
            get { return Contact(_subDirCommonSheets); }
        }

        public static string DirCommonVbt
        {
            get { return Contact(_subDirCommonVbt); }
        }

        public static string DirDcSpec
        {
            get { return Contact(_subDirDcSpec); }
        }

        public static string DirDevChar
        {
            get { return Contact(_subDirDevChar); }
        }

        public static string DirGlbSpec
        {
            get { return Contact(_subDirGlbSpec); }
        }

        public static string DirLevel
        {
            get { return Contact(_subDirLevel); }
        }

        public static string DirPinMap
        {
            get { return Contact(_subDirPinMap); }
        }

        public static string DirPorts
        {
            get { return Contact(_subDirPorts); }
        }

        public static string DirWaveDef
        {
            get { return Contact(_subDirAnalog); }
        }

        public static string DirMixedSignal
        {
            get { return Contact(_subDirAnalog); }
        }

        public static string DirTimings
        {
            get { return Contact(_subDirTimings); }
        }

        public static string DirJob
        {
            get { return Contact(_subDirJob); }
        }

        public static string DirPatSetsAll
        {
            get { return Contact(_subDirPatSetsAll); }
        }

        public static string DirEfuse
        {
            get { return Contact(_subDirEfuse); }
        }


        public static string DirScan
        {
            get
            {
                return Contact(_subDirScan);
            }
        }

        public static string DirSocScan
        {
            get
            {
                return Contact(_subDirSocScan);
            }
        }

        public static string DirSocMbist
        {
            get
            {
                return Contact(_subDirSocMbist);
            }
        }

        public static string DirCpuScan
        {
            get
            {
                return Contact(_subDirCpuScan);
            }
        }

        public static string DirMbist
        {
            get
            {
                return Contact(_subDirMbist);
            }
        }

        public static string DirCpuMbist
        {
            get
            {
                return Contact(_subDirCpuMbist);
            }
        }

        public static string DirGpuScan
        {
            get
            {
                return Contact(_subDirGpuScan);
            }
        }

        public static string DirGpuMbist
        {
            get
            {
                return Contact(_subDirGpuMbist);
            }
        }

        public static string DirHardIp
        {
            get { return Contact(_subDirHardIp); }
        }


        public static string DirLib
        {
            get { return Contact(_subDirLib); }
        }

        public static string DirConingLib
        {
            get { return Contact(_subDirCodingLib); }
        }

        public static string DirVbtGenTool
        {
            get { return Contact(_subDirVbtGenTool); }
        }

        public static string DirBinCut
        {
            get { return Contact(_subDirBinCut); }
        }

        public static string DirNonBinCut
        {
            get { return Contact(_subDirNonBinCut); }
        }

        public static string DirEvs
        {
            get { return Contact(_subDirEvs); }
        }

        public static string DirCpm
        {
            get { return Contact(_subDirCpm); }
        }

        public static string DirSpi
        {
            get
            {
                return Contact(_subDirSpi);
            }
        }

        public static string DirConti
        {
            get { return Contact(_subDirConti); }
        }

        public static string DirSpiRom
        {
            get
            {
                return Contact(_subDirSpiRom);
            }
        }

        public static string DirRtos
        {
            get
            {
                return Contact(_subDirRtos);
            }
        }

        public static string DirMain
        {
            get { return Contact(_subDirMain); }
        }

        public static string DirReference
        {
            get { return Contact(_subDirReference); }
        }

        public static string DirSubProgram
        {
            get { return Contact(_subDirSubProgram); }
        }

        public static string DirT0Tx
        {
            get { return Contact(_subDirT0Tx); }
        }
        #endregion
        public static string DirVreValid
        {
            get { return Contact(_subDirVreValid); }
        }
        internal static string Contact(string subDir)
        {
            return Path.Combine(TarDir, subDir);
        }

        public static void CreateFolder()
        {
            FolderReSetup();

            if (Directory.Exists(DirXmlFiles))
            {
                try
                {
                    Directory.Delete(DirXmlFiles, true);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }

            if (Directory.Exists(DirLogs))
            {
                try
                {
                    Directory.Delete(DirLogs, true);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }

            foreach (string dir in _dirList)
            {
                if (!Directory.Exists(dir))
                {
                    try
                    {
                        Directory.CreateDirectory(dir);
                    }
                    catch (Exception ex)
                    {
                        ErrorMessageBox.Show(string.Format(ex.ToString()));
                    }
                }
            }
        }

        private static void FolderReSetup()
        {
            _dirList = new List<string>
            {
                DirLogs,
                DirXmlFiles,
                DirAcSpec,
                DirBinTable,
                DirChannelMap,
                DirCommonSheets,
                DirCommonVbt,
                DirDcSpec,
                DirDevChar,
                DirGlbSpec,
                DirLevel,
                DirPinMap,
                DirPorts,
                DirTimings,
                DirJob,
                DirPatSetsAll,
                DirEfuse,
                DirScan,
                DirSocScan,
                DirSocMbist,
                DirCpuScan,
                DirMbist,
                DirCpuMbist,
                DirGpuScan,
                DirGpuMbist,
                DirHardIp,
                DirLib,
                DirConingLib,
                DirVbtGenTool,
                DirBinCut,
                DirNonBinCut,
                DirEvs,
                DirSpi,
                DirConti,
                DirSpiRom,
                DirRtos,
                DirMain,
                DirVreValid
            };
        }

        public static void Clear()
        {
            _dirList = null;
            _tarDir = null;
        }
    }
}
