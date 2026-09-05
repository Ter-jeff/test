using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using BinCutScriptLib.SetFunction.SetStartStep;
using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    [DebuggerDisplay("{Mode}")]
    public class InterpolationRow
    {
        public int Site = -1;
        public string Mode = string.Empty;
        public string LowMode = string.Empty;
        public string HighMode1 = string.Empty;
        public string HighMode2 = string.Empty;
        public double EnValue = -1;
        public int Eqidx = -1;
        public double SelectedLvcc;
        public EnLine EnLine = new();
        public bool SkipTest;
        public bool Bin4Cand;

        public void CheckeqNLinesSkip(StreamWriter streamWriter)
        {
            const string strSkip = "SKIPTEST";
            const string skipEqnLineKeyWord = "The lowest Performance Mode";
            string modeName = "";
            List<InterpolatioNode> interpolatioNotNodes = BinCutConfig.InterpolationNodes;

            if (!EnLine.Line.Contains(skipEqnLineKeyWord, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            List<string> spt = [.. EnLine.Line.Split([','], StringSplitOptions.RemoveEmptyEntries)];
            foreach (string tok in spt)
            {
                if (tok.Trim().StartsWithIgnoreCase("Binning Mode") && tok.Trim().Contains(':'))
                {
                    modeName = tok.Trim().Split([':'], StringSplitOptions.RemoveEmptyEntries)[1].Trim().ToUpper();
                }
                else if (tok.Trim().StartsWith("VDD_"))
                {
                    modeName = tok.Trim();
                }
            }

            string tf = spt.Find(x => x.ContainsIgnoreCase(strSkip))!.Split([':'], StringSplitOptions.RemoveEmptyEntries).Last().Trim();
            bool bSkip = tf.EqualsIgnoreCase("TRUE");
            bool expectSkip = interpolatioNotNodes.Find(x => x.Mode == modeName)!.Bskip;
            if (bSkip != expectSkip)
            {
                string msg = "The En skip setting of Interpolation mismatch";
                streamWriter.WriteLine(msg);
                msg = $"Fail line:{EnLine.LineNo} ";
                streamWriter.WriteLine(msg);
                msg = $"     Mode :{modeName}";
                streamWriter.WriteLine(msg);
                msg = $"     Expected :{expectSkip}";
                streamWriter.WriteLine(msg);
                msg = $"     Datalog  :{bSkip}";
                streamWriter.WriteLine(msg);
                msg = "";
                streamWriter.WriteLine(msg);
            }
        }

        public void CheckeqNLinesSkipCs(StreamWriter streamWriter)
        {
            string modeName = Mode;
            List<InterpolatioNode> interpolatioNotNodes = BinCutConfig.InterpolationNodes;
            string datalog = "True";
            string expectSkip = "False";
            InterpolatioNode? interpolation = interpolatioNotNodes.Find(x => x.Mode.Split('_').Last() == modeName);
            if (interpolation != null)
            {
                expectSkip = interpolation.Bskip.ToString();
            }

            bool fail = !datalog.EqualsIgnoreCase(expectSkip);
            if (fail || BinCutConfig.IsDebugPrint)
            {
                string msg = fail ? "Fail: The Interpolation mismatch" : "Pass: The Interpolation match";
                streamWriter.WriteLine(msg);
                msg = fail ? $"Fail line:{EnLine.LineNo} " : $"Pass line:{EnLine.LineNo} ";
                streamWriter.WriteLine(msg);
                msg = $"     Mode :{modeName}";
                streamWriter.WriteLine(msg);
                msg = $"     Expected :{expectSkip}";
                streamWriter.WriteLine(msg);
                msg = $"     Datalog  :{datalog}";
                streamWriter.WriteLine(msg);
                msg = "";
                streamWriter.WriteLine(msg);
            }
        }
    }
}
