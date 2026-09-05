namespace CommonLib.ErrorReport.Base
{
    public enum EnumErrorBehavior
    {
        //The same entity is defined more than once.​
        Duplicate = 1,
        //Data exists but does not meet expected rules or constraints.​
        Invalid = 2,
        //Conflicting definitions across related data.​
        Mismatch = 3,
        //Required data is absent or undefined.​
        Missing = 4,
        //Unnecessary or duplicated data with no practical value.​
        Redundant = 5,
        //Data breaks system rules or logical constraints.​
        RuleViolation = 6,
        //Info
        Info = 7,
    }
}
