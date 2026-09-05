namespace BinCutScriptLib.Base
{
    public class Alarm
    {
        internal Line.BinCutLineBase _alarmMessage = new();
        public bool IsBeforeBv;
        public string Type = string.Empty;
    }
}
