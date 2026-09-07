using System;
using System.Collections.Generic;

namespace Automation.Singleton
{
    public class IdsRepairSingleton
    {
        private static IdsRepairSingleton _instance;

        private readonly List<MbistRepairFlag> _mbistRepairFlags;
        private readonly List<IdsFlag> _idsFlags;

        private IdsRepairSingleton()
        {
            _mbistRepairFlags = new List<MbistRepairFlag>();
            _idsFlags = new List<IdsFlag>();
        }

        public static IdsRepairSingleton Instance()
        {
            return _instance ?? (_instance = new IdsRepairSingleton());
        }

        public static void Initialize()
        {
            _instance = null;
        }

        public bool AddIdsFlag(string flagName, string n)
        {
            if (_idsFlags.Exists(p => p.Nx.Equals(n)))
            {
                return false;
            }
            _idsFlags.Add(new IdsFlag(flagName, n));

            return true;
        }

        public bool AddMbistRepairFlag(string module, string flag)
        {
            if (_mbistRepairFlags.Exists(p => p.Module.Equals(module, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }
            _mbistRepairFlags.Add(new MbistRepairFlag(module, flag));
            return true;
        }
    }

    public class MbistRepairFlag
    {
        public string Module;
        public string Flag;

        public MbistRepairFlag(string module, string flag)
        {
            Module = module;
            Flag = flag;
        }
    }

    public class IdsFlag
    {
        public string Flag;
        public string Nx;

        public IdsFlag(string flagName, string number)
        {
            Flag = flagName;
            Nx = number;
        }
    }
}
