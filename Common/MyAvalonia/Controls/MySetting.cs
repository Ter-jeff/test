using System;
using System.Collections.Generic;

using MyAvalonia.Model;

namespace MyAvalonia.Controls
{
    public class MySetting : IMySetting
    {
        public string ClassType { get; set; } = string.Empty;
        public List<MyControl> MyControls { get; set; } = [];

        public string GetValue(string name)
        {
            foreach (MyControl myControl in MyControls)
            {
                if (myControl.ControlName.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return myControl.Content;
                }
            }
            return string.Empty;
        }
    }
}
