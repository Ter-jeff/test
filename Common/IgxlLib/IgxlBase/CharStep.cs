namespace IgxlLib.IgxlBase
{
    public class CharStep : IgxlRow
    {
        public string VoltageType { get; set; } = string.Empty;
        public string SetupName { get; set; } = string.Empty;
        public string StepName { get; set; } = string.Empty;
        public string Mode { get; set; } = string.Empty;
        //Parameter
        public string ParameterType { get; set; } = string.Empty;
        public string ParameterName { get; set; } = string.Empty;
        //Range
        public string RangeCalcField { get; set; } = string.Empty;
        public string RangeFrom { get; set; } = string.Empty;
        public string RangeTo { get; set; } = string.Empty;
        public string RangeSteps { get; set; } = string.Empty;
        public string RangeStepSize { get; set; } = string.Empty;
        //PerformTest
        public string PerformTest { get; set; } = string.Empty;
        //Test Limits
        public string TestLimitLow { get; set; } = string.Empty;
        public string TestLimitHigh { get; set; } = string.Empty;
        //Algorithm
        public string AlgorithmName { get; set; } = string.Empty;
        public string AlgorithmArguments { get; set; } = string.Empty;
        public string AlgorithmResultsCheck { get; set; } = string.Empty;
        public string AlgorithmTransition { get; set; } = string.Empty;
        //Apply To
        public string ApplyToPinExecMode { get; set; } = string.Empty;
        public string ApplyToPins { get; set; } = string.Empty;
        public string ApplyToTimeSets { get; set; } = string.Empty;
        //Device Margin
        public string DeviceMarginContexts { get; set; } = string.Empty;
        public string DeviceMarginInstances { get; set; } = string.Empty;
        public string DeviceMarginPatterns { get; set; } = string.Empty;
        //Adjust
        public string AdjustBackoff { get; set; } = string.Empty;
        public string AdjustFromSetup { get; set; } = string.Empty;
        public string AdjustSpecName { get; set; } = string.Empty;
        //Function
        public string AxisExecutionOrder { get; set; } = string.Empty;
        public string Function { get; set; } = string.Empty;
        //Arguments
        public string Arguments { get; set; } = string.Empty;
        public string InterposeFunctions { get; set; } = string.Empty;
        //Interpose Functions
        public string PreSetup { get; set; } = string.Empty;
        public string PreSetupArguments { get; set; } = string.Empty;
        public string PreStep { get; set; } = string.Empty;
        public string PreStepArguments { get; set; } = string.Empty;
        public string PrePoint { get; set; } = string.Empty;
        public string PrePointArguments { get; set; } = string.Empty;
        public string PostPoint { get; set; } = string.Empty;
        public string PostPointArguments { get; set; } = string.Empty;
        public string PostStep { get; set; } = string.Empty;
        public string PostStepArguments { get; set; } = string.Empty;
        public string PostSetup { get; set; } = string.Empty;
        public string PostSetupArguments { get; set; } = string.Empty;
        //Output
        public string OutputFormat { get; set; } = string.Empty;
        public string OutputTextFile { get; set; } = string.Empty;
        public string OutputSheet { get; set; } = string.Empty;
        public string OutputSuspendDatalog { get; set; } = string.Empty;
        //Output Destinations
        public string OutputDestinationsTextFile { get; set; } = string.Empty;
        public string OutputDestinationsSheet { get; set; } = string.Empty;
        public string OutputDestinationsDatalog { get; set; } = string.Empty;
        public string OutputDestinationsImmediateWin { get; set; } = string.Empty;
        public string OutputDestinationsOutputWin { get; set; } = string.Empty;
        //Comment
        public string Comment { get; set; } = string.Empty;

        public CharStep(string setupName, string stepName)
        {
            SetupName = setupName;
            StepName = stepName;
            AxisExecutionOrder = "X-Y[-Z]";
            OutputSuspendDatalog = "TRUE";
        }

        public CharStep(CharStep charStep) : base(charStep)
        {
            if (charStep == null)
            {
                return;
            }

            VoltageType = charStep.VoltageType;
            SetupName = charStep.SetupName;
            StepName = charStep.StepName;
            Mode = charStep.Mode;
            ParameterType = charStep.ParameterType;
            ParameterName = charStep.ParameterName;
            RangeCalcField = charStep.RangeCalcField;
            RangeFrom = charStep.RangeFrom;
            RangeTo = charStep.RangeTo;
            RangeSteps = charStep.RangeSteps;
            RangeStepSize = charStep.RangeStepSize;
            PerformTest = charStep.PerformTest;
            TestLimitLow = charStep.TestLimitLow;
            TestLimitHigh = charStep.TestLimitHigh;
            AlgorithmName = charStep.AlgorithmName;
            AlgorithmArguments = charStep.AlgorithmArguments;
            AlgorithmResultsCheck = charStep.AlgorithmResultsCheck;
            AlgorithmTransition = charStep.AlgorithmTransition;
            ApplyToPinExecMode = charStep.ApplyToPinExecMode;
            ApplyToPins = charStep.ApplyToPins;
            ApplyToTimeSets = charStep.ApplyToTimeSets;
            DeviceMarginContexts = charStep.DeviceMarginContexts;
            DeviceMarginInstances = charStep.DeviceMarginInstances;
            DeviceMarginPatterns = charStep.DeviceMarginPatterns;
            AdjustBackoff = charStep.AdjustBackoff;
            AdjustFromSetup = charStep.AdjustFromSetup;
            AdjustSpecName = charStep.AdjustSpecName;
            AxisExecutionOrder = charStep.AxisExecutionOrder;
            Function = charStep.Function;
            Arguments = charStep.Arguments;
            InterposeFunctions = charStep.InterposeFunctions;
            PreSetup = charStep.PreSetup;
            PreSetupArguments = charStep.PreSetupArguments;
            PreStep = charStep.PreStep;
            PreStepArguments = charStep.PreStepArguments;
            PrePoint = charStep.PrePoint;
            PrePointArguments = charStep.PrePointArguments;
            PostPoint = charStep.PostPoint;
            PostPointArguments = charStep.PostPointArguments;
            PostStep = charStep.PostStep;
            PostStepArguments = charStep.PostStepArguments;
            PostSetup = charStep.PostSetup;
            PostSetupArguments = charStep.PostSetupArguments;
            OutputFormat = charStep.OutputFormat;
            OutputTextFile = charStep.OutputTextFile;
            OutputSheet = charStep.OutputSheet;
            OutputSuspendDatalog = charStep.OutputSuspendDatalog;
            OutputDestinationsTextFile = charStep.OutputDestinationsTextFile;
            OutputDestinationsSheet = charStep.OutputDestinationsSheet;
            OutputDestinationsDatalog = charStep.OutputDestinationsDatalog;
            OutputDestinationsImmediateWin = charStep.OutputDestinationsImmediateWin;
            OutputDestinationsOutputWin = charStep.OutputDestinationsOutputWin;
            Comment = charStep.Comment;
        }

        public CharStep Copy()
        {
            return new CharStep(this);
        }
    }
}
