using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public class EFuseStdfReader1(string filename, List<string> hipList)
    {
        private readonly string _headerX = "DIE_X";
        private readonly string _headerY = "DIE_Y";
        private readonly string _headerHighL = "HighL";
        private readonly string _headerLowL = "LowL";
        private readonly string _headerPid = "PID";
        private readonly List<int> _pidIndx = [];
        private readonly string _headerParameter = "Parameter";
        private readonly List<string> _hipItem = hipList;
        private readonly string _filename = filename;

        private int _highLIndx = -1;
        private int _lowLIndx = -1;
        private int _parameterIndx = -1;
        public Dictionary<string, string> LowLDic = [];
        public Dictionary<string, string> HighLDic = [];
        public List<HipItem> HipItems = [];
        public List<PrrItem> PrRs = [];
        private bool _prRflag;

        public void WorkFlow()
        {
            if (string.IsNullOrEmpty(_filename))
            {
                return;
            }

            var reader = new StreamReader(_filename);
            string? line = "";
            var prrItem = new PrrItem();
            while ((line = reader.ReadLine()) != null)
            {
                if (line.Contains(_headerParameter) &&
                    line.Contains(_headerHighL) &&
                    line.Contains(_headerLowL) &&
                    line.Contains(_headerPid))
                {
                    string[] linespt = line.Split(',');
                    for (int i = 0; i < linespt.Length; i++)
                    {
                        if (linespt[i].EqualsIgnoreCase(_headerParameter))
                        {
                            _parameterIndx = i;
                        }

                        if (linespt[i].EqualsIgnoreCase(_headerHighL))
                        {
                            _highLIndx = i;
                        }

                        if (linespt[i].EqualsIgnoreCase(_headerLowL))
                        {
                            _lowLIndx = i;
                        }

                        if (linespt[i].Contains(_headerPid))
                        {
                            _pidIndx.Add(i);
                            var item = new HipItem(linespt[i]);
                            HipItems.Add(item);
                        }
                    }
                }
                if (line.Contains(_headerX))
                {
                    string[] linespt = line.Split(',');
                    int k = 0;
                    foreach (HipItem item in HipItems)
                    {
                        item.X = linespt[_pidIndx[k]];
                        k++;
                    }
                }
                if (line.Contains(_headerY))
                {
                    string[] linespt = line.Split(',');
                    int k = 0;
                    foreach (HipItem item in HipItems)
                    {
                        item.Y = linespt[_pidIndx[k]];
                        k++;
                    }
                }

                if (_parameterIndx != -1)
                {
                    if (_hipItem.FirstOrDefault(p => line.ContainsIgnoreCase(p.ToLower())) != null)
                    {
                        string[] linespt = line.Split(',');
                        string name = _hipItem.FirstOrDefault(p => line.ContainsIgnoreCase(p.ToLower()))!;
                        string limitL = linespt[_lowLIndx];
                        string limitH = linespt[_highLIndx];
                        LowLDic.Add(name, limitL);
                        HighLDic.Add(name, limitH);
                        int k = 0;
                        foreach (HipItem item in HipItems)
                        {
                            item.HipData.Add(name, linespt[_pidIndx[k]]);
                            k++;
                        }
                    }
                }
                if (_prRflag && !line.Contains(':'))
                {
                    prrItem.PrrCode = line.Trim();
                    _prRflag = false;
                }
                else
                {
                    _prRflag = false;
                }
                if (line.ContainsIgnoreCase("PRR:"))
                {

                    //PRR:1|0|25|17279|P|1|1|3|16|||155570|0003788A1180013A
                    List<string> prrInfos = [.. line.Split(':')[1].Trim().Split('|')];
                    prrItem = new PrrItem { X = prrInfos[7], Y = prrInfos[8], PrrCode = prrInfos.Last() };
                    _prRflag = string.IsNullOrEmpty(prrItem.PrrCode);
                    PrRs.Add(prrItem);
                }
            }
        }
    }
}
