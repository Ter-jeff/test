using CommonLib.Extension;
namespace TestPlanLib.BinCut
{
    public class Domain
    {
        public string Name = "";

        public new bool Equals(object obj)
        {
            string newName = ((Domain)obj).Name;
            if (Name == newName)
            {
                return true;
            }

            if (Name.EqualsIgnoreCase("GFX") && newName.EqualsIgnoreCase("GPU"))
            {
                return true;
            }

            return Name.EqualsIgnoreCase("GPU") && newName.EqualsIgnoreCase("GFX");
        }
    }
}
