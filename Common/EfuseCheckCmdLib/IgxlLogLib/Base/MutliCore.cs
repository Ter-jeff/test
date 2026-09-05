using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class MutliCore
    {
        private readonly string _targerBlock = "(SA|TD)(CHAIN)*";
        private readonly string _passKeyWord = "P";

        private int _coreCnt;
        public List<CorePassFail> CoreResult = [];

        public void SetData(string block, string subBlock, string data)
        {
            CoreResult.Add(new CorePassFail(block, subBlock, data));
            _coreCnt = data.Split(',').Length;
        }

        public int GetCoreByPassIndex(string module)
        {

            var corePassList = new List<bool>();
            for (int i = 0; i < _coreCnt; i++)
            {
                corePassList.Add(true);
            }

            List<CorePassFail> targetModuelCoreResult =
                CoreResult.FindAll(p => Regex.IsMatch(p.Block, module + _targerBlock, RegexOptions.IgnoreCase) && Regex.IsMatch(p.SubBlock, _targerBlock, RegexOptions.IgnoreCase));

            if (targetModuelCoreResult.Count == 0)
            {
                // no block core info => allcorepass -1
                return -1;
            }

            foreach (
                CorePassFail targetItem in targetModuelCoreResult)
            {
                string[] pfAry = targetItem.Data.Split(',');

                for (int i = 0; i < pfAry.Length; i++)
                {
                    if (pfAry[i] != _passKeyWord)
                    {
                        corePassList[i] = false;
                    }
                }
            }

            if (corePassList.FindAll(p => !p).Count > 1)
            {
                // two core fail , should not test 
                return -1;
            }

            for (int i = 0; i < corePassList.Count; i++)
            {
                if (!corePassList[i])
                {
                    return i;
                }
            }
            // all core pass , just bypass core 0
            return 0;
        }
    }
}
