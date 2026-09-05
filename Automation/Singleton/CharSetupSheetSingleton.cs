using System.Collections.Generic;

using IgxlLib.IgxlSheets;

namespace Automation.Singleton
{
    public class CharSetupSheetSingleton
    {
        private static CharSetupSheetSingleton _instance;
        private static readonly object _locker = new object();
        public const string CharSetup1DSheetName = "CharSetUp_1D";

        public static CharSetupSheetSingleton Instance()
        {
            if (_instance == null)
            {
                lock (_locker)
                {
                    if (_instance == null)
                    {
                        _instance = new CharSetupSheetSingleton();
                    }
                }
            }
            return _instance;
        }

        public static void Initialize()
        {
            _instance = new CharSetupSheetSingleton { CharSheets = new List<CharSheet>() };
        }

        private CharSetupSheetSingleton()
        {
            CharSheets = new List<CharSheet>();
        }

        public List<CharSheet> CharSheets { get; private set; }

        public CharSheet Get1DCharSheet()
        {
            CharSheet charSheet = CharSheets.Find(p => p.Name.Equals(CharSetup1DSheetName));
            if (charSheet == null)
            {
                charSheet = new CharSheet(CharSetup1DSheetName);
                CharSheets.Add(charSheet);
            }
            return charSheet;
        }
    }
}
