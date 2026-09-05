using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TerLintAnalyzer.Ter.Rules;
using TerLintAnalyzer.Ter.StaticSettable;
using TerLintAnalyzer.Test.Verifiers;

namespace TerLintAnalyzer.Test.UT
{
    [TestClass]
    public class StaticSettableCodeFixProviderTests : CSharpCodeFixVerifier<StaticSettableAnalyzer, StaticSettableCodeFixProvider>
    {
        [TestMethod]
        public async Task Field_NoInitializer_ConvertedToAutoPropertyWithPrivateSetter()
        {
            string test = $@"
public class Config
{{
    public static int {{|{TerRules.StaticSettableRule.Id}:MyField|}};
}}";

            string fixedCode = @"
public class Config
{
    public static int MyField { get; private set; }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Field_WithInitializer_ConvertedToAutoPropertyWithPrivateSetter()
        {
            string test = $@"
public class Config
{{
    public static string {{|{TerRules.StaticSettableRule.Id}:Name = ""abc""|}};
}}";

            string fixedCode = @"
public class Config
{
    public static string Name { get; private set; } = ""abc"";
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }

        [TestMethod]
        public async Task Property_AutoSetter_MadePrivate()
        {
            string test = $@"
public class Config
{{
    public static int {{|{TerRules.StaticSettableRule.Id}:MyProperty|}} {{ get; set; }}
}}";

            string fixedCode = @"
public class Config
{
    public static int MyProperty { get; private set; }
}";

            await VerifyCodeFixAsync(test, fixedCode);
        }
    }
}
