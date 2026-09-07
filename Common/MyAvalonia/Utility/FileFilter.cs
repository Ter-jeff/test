using MyAvalonia.Model;

namespace MyAvalonia.Utility
{
    public static class FileFilter
    {
        public static string GetFilter(EnumFileFilter fileFilter) => fileFilter switch
        {
            EnumFileFilter.TestProgram => "Test Program Files (*.igxl;*.xls*)|*.igxl;*.xls*",
            EnumFileFilter.Igxl => "IGXL Files (*.igxl)|*.igxl",
            EnumFileFilter.PatternCsv => "Pattern CSV Files (*pattern*.csv)|*pattern*.csv",
            EnumFileFilter.Excel => "Excel Workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            EnumFileFilter.IdsDistribution => "IDS Distribution (IDS_Distribution.txt)|IDS_Distribution.txt*",
            EnumFileFilter.TestFlowSheet => "Test Flow Sheet (*test*plan*.xlsx*)|*test*plan*.xlsx*",
            EnumFileFilter.BinCut => "Bin Cut Files (*.txt;*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx",
            EnumFileFilter.BinCutPost => "Post Bin Cut Files (*.txt;*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx",
            EnumFileFilter.TemplateFile => "Template Files (*.tmp)|*.tmp",
            EnumFileFilter.Txt => "Text Files (*.txt)|*.txt",
            EnumFileFilter.DataLog => "Data Log Files (*.txt;*.txt.gz)|*.txt;*.txt.gz",
            EnumFileFilter.BasFile => "BAS Files (*.bas)|*.bas",
            EnumFileFilter.PaFile => "PA Files (*.csv;*.xls*)|*.csv;*.xlsx;*.xlsm",
            EnumFileFilter.XmlFile => "XML Files (*.xml)|*.xml",
            EnumFileFilter.Csv => "CSV Files (*.csv)|*.csv",
            EnumFileFilter.YamlFile => "YAML Files (*.yaml)|*.yaml",
            EnumFileFilter.ExcelOld => "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls",
            EnumFileFilter.OtpFile => "OTP Files (*.otp)|*.otp",
            EnumFileFilter.StdfFile => "STDF Files (*.std;*.std.gz)|*.std;*.std.gz",
            EnumFileFilter.TestPlan => "Test Plan (*test*plan*.xlsx*)|*test*plan*.xlsx*",
            EnumFileFilter.BdFile => "BD Files (*.csv)|*.csv",
            EnumFileFilter.Atp => "ATP Files (*.atp)|*.atp",
            EnumFileFilter.HardipInfo => "Hardip Info Files (*.txt;*.log)|*.txt;*.log",
            EnumFileFilter.DDR => "Datalog DDR ELB/ILB (*.*)|*.*",
            EnumFileFilter.Pm => "PM Files (*.pm)|*.pm",
            _ => string.Empty
        };
    }
}
