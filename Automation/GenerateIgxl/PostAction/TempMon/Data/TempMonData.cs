using System;

using Automation.GenerateIgxl.PostAction.TempMon.Enums;

namespace Automation.GenerateIgxl.PostAction.TempMon.Data
{
    public class TempMonData
    {
        public string Mode { get; set; }
        public EnumCondition Condition { get; set; }
        public EnumType Type { get; set; }
        public string Item { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is TempMonData other)
            {
                return string.Equals(Mode, other.Mode, StringComparison.OrdinalIgnoreCase)
                    && Condition == other.Condition
                    && Type == other.Type
                    && string.Equals(Item, other.Item, StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        public override int GetHashCode()
        {
            int hashMode = Mode?.ToLowerInvariant().GetHashCode() ?? 0;
            int hashItem = Item?.ToLowerInvariant().GetHashCode() ?? 0;

            return hashMode ^ Condition.GetHashCode() ^ Type.GetHashCode() ^ hashItem;
        }
    }
}
