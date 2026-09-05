using System;
using System.Drawing;

using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    public class EnLine : BinCutLineBase
    {
        public InterpolationRow GetEnRow()
        {
            //case 1
            //STEP1b. Get the p.power that should apply the inner difference method
            //Site:1, CurrentPassBinCutNum:2, VDD_PCPU_MP101, The En Voltage: 659.375 mV. The lowest Performance Mode: VDD_PCPU_MP001, The highest Performance Mode: VDD_PCPU_MP002, The MFx: 0, SkipTest: True
            //Site:1, CurrentPassBinCutNum:2, VDD_PCPU_MP102, The En Voltage: 693.75 mV. The lowest Performance Mode: VDD_PCPU_MP001, The highest Performance Mode: VDD_PCPU_MP002, The MFx: 0, SkipTest: True
            //case 2
            //Site:0,CurrentPassBinCutNum:1,Binning Mode:VDD_GPU_MG003,The lowest Performance Mode:VDD_GPU_MG001, The highest Performance Mode:VDD_GPU_MG004, The MFx:0.542857142857143,SkipTest:False,The interpolated Voltage:593.75 mV,The selected Eqn:0,The selected Voltage:-1
            //Site:2,CurrentPassBinCutNum:1,Binning Mode:VDD_GPU_MG003,The lowest Performance Mode:VDD_GPU_MG001, The highest Performance Mode:VDD_GPU_MG004, The MFx:0.542857142857143,SkipTest:False,The interpolated Voltage:600 mV,The selected Eqn:0,The selected Voltage:-1
            string[] spt = Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            var enRow = new InterpolationRow();
            foreach (string tok in spt)
            {
                if (tok.Trim().StartsWithIgnoreCase("Binning Mode") &&
                    tok.Trim().Contains(':'))
                {
                    enRow.Mode = tok.Split([':'])[1].Trim();
                }

                else if (tok.Trim().StartsWith("VDD_"))
                {
                    enRow.Mode = tok.Trim();
                }
                else if (tok.Trim().Contains("THE LOWEST PERFORMANCE POWER", StringComparison.OrdinalIgnoreCase) ||
                         tok.Trim().Contains("THE LOWEST PERFORMANCE MODE", StringComparison.OrdinalIgnoreCase))
                {
                    enRow.LowMode = tok.Split([':'])[1].Trim();
                }
                else if (tok.Trim().Contains("THE HIGHEST PERFORMANCE POWER", StringComparison.OrdinalIgnoreCase) ||
                         tok.Trim().Contains("THE HIGHEST PERFORMANCE MODE", StringComparison.OrdinalIgnoreCase)) //Jerry has an TYPO here
                {
                    enRow.HighMode1 = tok.Trim().Split([':'])[1].Trim();
                    if (enRow.HighMode1.Contains('/'))
                    {
                        string[] tmp = enRow.HighMode1.Split('/');
                        enRow.HighMode1 = tmp[0].Trim();
                        enRow.HighMode2 = tmp[1].Trim();
                    }
                }
                else if (tok.Trim().Contains("THE EN VOLTAGE", StringComparison.OrdinalIgnoreCase) ||
                         tok.Trim().Contains("THE INTERPOLATED VOLTAGE", StringComparison.OrdinalIgnoreCase))
                {
                    string value = Reg.RegexVoltage.Match(tok.Trim()).Groups["value"].ToString();
                    if (!double.TryParse(value, out double enValue))
                    {
                        BinCutController.Controller.RichTextBoxAppend($"stream txt: {tok} try Parse fail", Color.Red);
                    }

                    enRow.EnValue = enValue;
                }
                else if (tok.Trim().StartsWithIgnoreCase("SITE:"))
                {
                    enRow.Site = int.Parse(tok.Trim().Split(':')[1]);
                }
                else if (tok.Trim().StartsWithIgnoreCase("SkipTest"))
                {
                    enRow.SkipTest = tok.Trim().Split(':')[1].EqualsIgnoreCase("True");
                }
                else if (tok.Trim().StartsWithIgnoreCase("The selected Eqn") && tok.Trim().Contains(':'))
                {
                    string strEqidx = tok.Trim().Split([':'], StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                    if (!int.TryParse(strEqidx, out int eqidx))
                    {
                        BinCutController.Controller.RichTextBoxAppend($"stream txt: {tok} try Parse fail", Color.Red);
                    }

                    enRow.Eqidx = eqidx;
                }
                if (tok.Trim().StartsWithIgnoreCase("The selected Voltage"))
                {
                    string value = Reg.RegexVoltage.Match(tok.Trim()).Groups["value"].ToString();
                    if (!double.TryParse(value, out double logStepLvcc))
                    {
                        BinCutController.Controller.RichTextBoxAppend($"stream txt: {tok} try Parse fail", Color.Red);
                    }

                    enRow.SelectedLvcc = logStepLvcc;
                }
                enRow.EnLine = this;
            }
            return enRow;
        }
    }
}
