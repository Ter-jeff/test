namespace RfLib.Dvdc.Base
{
    public class Component(string name, string type)
    {
        public string Name = name;
        public string Type = type;
        public string Status = "";

        public Component Clone()
        {
            var newItem = new Component(Name, Type) { Status = Status };
            return newItem;
        }

    }
}
