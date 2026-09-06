using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MyAvalonia.Model
{
    public class HelpFileRow : INotifyPropertyChanged
    {
        private bool _select;

        public bool Select
        {
            get => _select;
            set
            {
                if (_select != value)
                {
                    _select = value;
                    OnPropertyChanged();
                }
            }
        }

        public string FileName { get; internal set; } = string.Empty;
        public string GroupName { get; internal set; } = string.Empty;
        public string Key { get; internal set; } = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public record HelpGroupItem(string Name);
}
