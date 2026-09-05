using System.Collections.Generic;
using System.Linq;

namespace BinCutScriptLib.Reader
{
    public class InheritanceManager
    {
        private List<Dictionary<string, bool>> _inheritStatus = [];

        public void Read(List<Dictionary<string, bool>> binningTb)
        {
            _inheritStatus = binningTb;
        }

        public List<List<string>> GetEnableInheritLists()
        {
            return [.. _inheritStatus.Select(inheritList => inheritList.ToList().FindAll(x => x.Value).Select(x => x.Key).ToList())];
        }

        public List<List<string>> GetAllInheritLists()
        {
            return [.. _inheritStatus.Select(inheritList => inheritList.ToList().Select(x => x.Key).ToList())];
        }

        public void SetInheritModeEnable(string mode)
        {
            foreach (Dictionary<string, bool> inheritList in _inheritStatus.Where(inheritList => inheritList.ContainsKey(mode)))
            {
                inheritList[mode] = true;
            }
        }
    }
}
