using System.Collections.Generic;
using System.Linq;

namespace IgxlLib.IgxlBase
{
    public class CharSetup : IgxlRow
    {
        public string SetupName { get; set; } = string.Empty;
        public string TestMethod { get; set; } = string.Empty;
        public List<CharStep> CharSteps { get; set; } = [];

        public CharSetup()
        {
        }

        public CharSetup(CharSetup charSetup) : base(charSetup)
        {
            if (charSetup == null)
            {
                return;
            }

            SetupName = charSetup.SetupName;
            TestMethod = charSetup.TestMethod;

            // Deep copy the list of CharStep objects
            // This assumes CharStep has a Copy() method or a Copy Constructor
            CharSteps = charSetup.CharSteps?.Select(step => step.Copy()).ToList() ?? [];
        }

        public void AddStep(CharStep charStep)
        {
            CharSteps.Add(charStep);
        }
    }
}
