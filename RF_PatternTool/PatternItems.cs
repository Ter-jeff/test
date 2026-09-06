using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RF_PatternTool
{
    public class PatternItem : INotifyPropertyChanged
    {
        private string _status;
        private bool _isGenerate;
        public bool IsGenerate
        {
            get { return _isGenerate; }
            set
            {
                _isGenerate = value;
                OnPropertyChanged();
            }
        }
        private bool _isOverWrite;
        public bool IsOverWrite
        {
            get { return _isOverWrite; }
            set
            {
                _isOverWrite = value;
                OnPropertyChanged();
            }
        }

        public string Status
        {
            get { return _status; }
            set
            {
                _status = value;
                OnPropertyChanged();
            }
        }

        public string BenchLog { get; set; }
        public string Name { get; set; }
        public string PatName { get; set; }
        public string FilePath { get; set; }
        public string Address { get; set; }
        public int Index { get; set; }
        public List<string> Init { get; set; }
        public PatternItem()
        {
            IsGenerate = true;
            IsOverWrite = true;
            Status = "";
            BenchLog = "";
            Name = "";
            PatName = "";
            Address = "";
            Index = -1;
            Init = new List<string>();
        }

        protected void OnPropertyChanged([CallerMemberName] string property = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

    }
}
