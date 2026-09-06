using System;
using System.Diagnostics.CodeAnalysis;

using Avalonia;
using Avalonia.Controls;

namespace MyAvalonia.Controls
{
    [SuppressMessage("Maintainability", "CA1501:Avoid excessive inheritance", Justification = "Avalonia UI control inheritance depth")]
    public partial class MyStatusBar : UserControl
    {
        private DateTime _startTime;

        public static readonly StyledProperty<string?> StatusTextProperty =
            AvaloniaProperty.Register<MyStatusBar, string?>(nameof(StatusText));

        public string? StatusText
        {
            get => GetValue(StatusTextProperty);
            set => SetValue(StatusTextProperty, value);
        }

        public static readonly StyledProperty<int> ProgressBarValueProperty =
            AvaloniaProperty.Register<MyStatusBar, int>(nameof(ProgressBarValue));

        public int ProgressBarValue
        {
            get => GetValue(ProgressBarValueProperty);
            set => SetValue(ProgressBarValueProperty, value);
        }

        public static readonly StyledProperty<string?> ProcessTimeTextProperty =
            AvaloniaProperty.Register<MyStatusBar, string?>(nameof(ProcessTimeText));

        public string? ProcessTimeText
        {
            get => GetValue(ProcessTimeTextProperty);
            set => SetValue(ProcessTimeTextProperty, value);
        }

        public MyStatusBar()
        {
            InitializeComponent();
            DataContext = this;
        }

        public void TimeStart()
        {
            _startTime = TimeProvider.System.GetLocalNow().DateTime;
            ProgressBarValue = 0;
            ProcessTimeText = string.Empty;
            StatusText = string.Empty;
        }

        public void TimeStop()
        {
            ProgressBarValue = 0;
            ProcessTimeText = $"Process time : {TimeProvider.System.GetLocalNow().DateTime - _startTime:hh\\:mm\\:ss}";
            StatusText = "Done";
        }
    }
}
