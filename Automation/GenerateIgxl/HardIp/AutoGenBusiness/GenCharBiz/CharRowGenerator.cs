using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenCharBiz
{
    public class CharRowGenerator
    {
        public List<CharSetup> GenCharRow(List<HardIpPattern> planDic)
        {
            List<CharSetup> setup = new List<CharSetup>();

            foreach (HardIpPattern pattern in planDic)
            {
                if (pattern.Shmoo.CharSteps.Any())
                {
                    foreach (CharSetup charsetup in HardipCharSetup.GetShmoo(pattern).OfType<CharSetup>().ToList())
                    {
                        if (!setup.Exists(x => x.SetupName.Equals(charsetup.SetupName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            setup.Add(charsetup);
                        }
                    }
                }
            }
            return setup;
        }
    }
}
