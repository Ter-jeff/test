namespace TestPlanLib.Singleton
{
    public sealed class ModuleSingleton
    {
        #region Singleton

        // Fields
        private static ModuleSingleton? _instance;

        public string ModuleCpu;
        public string ModuleGfx;
        public string ModuleSoc;
        public string ModuleRtos;
        public string ModuleDdr;

        // Constructor
        private ModuleSingleton()
        {
            ModuleCpu = "Cpu";
            ModuleGfx = "Gfx";
            ModuleSoc = "Soc";
            ModuleRtos = "Rtos";
            ModuleDdr = "DDR";
        }

        // Methods
        public static ModuleSingleton Instance()
        {
            return _instance ??= new ModuleSingleton();
        }

        #endregion
        public static string GetModuleByPerformanceMode(string performanceMode)
        {
            if (performanceMode.Length < 5)
            {
                return "";
            }

            string moduleKeyWord = performanceMode.Substring(1, 1);
            return moduleKeyWord switch
            {
                //PCPU
                "C" => Instance().ModuleCpu,
                //PCPU
                "P" => Instance().ModuleCpu,
                //ECPU
                "E" => Instance().ModuleCpu,
                //ANE
                "A" => Instance().ModuleSoc,
                //GPU
                "G" => Instance().ModuleGfx,
                //SOC
                "S" => Instance().ModuleSoc,
                //DISP
                "I" => Instance().ModuleSoc,
                //DDR
                "D" => Instance().ModuleDdr,
                _ => "",
            };
        }
    }
}
