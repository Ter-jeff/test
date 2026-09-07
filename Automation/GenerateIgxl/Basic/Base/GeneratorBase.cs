namespace Automation.GenerateIgxl.Basic.Base
{
    public class GeneratorBase
    {
        protected const string PinNamePrefix = "_";
        protected const string FormulaPrefix = "=";
        protected const char FormulaConnector = '*';
        protected const string GlobalSpecSuffix = "GLB";

        protected string JoinPinComponent(string connector, params string[] components)
        {
            return string.Join(connector, components);
        }
    }
}
