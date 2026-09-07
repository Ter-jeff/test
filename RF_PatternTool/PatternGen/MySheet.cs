namespace TestPlanLib
{
    public abstract class MySheet
    {
        public string SheetName;


        public MySheet DimensionError()
        {
            return this;
        }

        public MySheet FirstHeaderError(string firstHeader)
        {
            return this;
        }
    }
}
