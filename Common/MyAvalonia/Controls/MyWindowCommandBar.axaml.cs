using System;
using System.Diagnostics.CodeAnalysis;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace MyAvalonia.Controls
{
    [SuppressMessage("Maintainability", "CA1501:Avoid excessive inheritance", Justification = "Avalonia UI control inheritance depth")]
    public partial class MyWindowCommandBar : UserControl
    {
        public static readonly StyledProperty<bool> ShowHelpProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, bool>(nameof(ShowHelp), true);

        public static readonly StyledProperty<bool> ShowIconProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, bool>(nameof(ShowIcon), true);

        public static readonly StyledProperty<string?> TitleProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, string?>(nameof(Title));

        public static readonly StyledProperty<HorizontalAlignment> TitleHorizontalAlignmentProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, HorizontalAlignment>(nameof(TitleHorizontalAlignment), HorizontalAlignment.Left);

        public static readonly StyledProperty<Control?> SettingBarProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, Control?>(nameof(SettingBar));

        public static readonly StyledProperty<IImage?> IconProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, IImage?>(nameof(Icon));

        public static readonly StyledProperty<IBrush?> TitleBackgroundProperty =
            AvaloniaProperty.Register<MyWindowCommandBar, IBrush?>(nameof(TitleBackground));

        public bool ShowHelp
        {
            get => GetValue(ShowHelpProperty);
            set => SetValue(ShowHelpProperty, value);
        }

        public bool ShowIcon
        {
            get => GetValue(ShowIconProperty);
            set => SetValue(ShowIconProperty, value);
        }

        public string? Title
        {
            get => GetValue(TitleProperty);
            set => SetValue(TitleProperty, value);
        }

        public HorizontalAlignment TitleHorizontalAlignment
        {
            get => GetValue(TitleHorizontalAlignmentProperty);
            set => SetValue(TitleHorizontalAlignmentProperty, value);
        }

        public Control? SettingBar
        {
            get => GetValue(SettingBarProperty);
            set => SetValue(SettingBarProperty, value);
        }

        public IImage? Icon
        {
            get => GetValue(IconProperty);
            set => SetValue(IconProperty, value);
        }

        public IBrush? TitleBackground
        {
            get => GetValue(TitleBackgroundProperty);
            set => SetValue(TitleBackgroundProperty, value);
        }

        public event EventHandler? HelpClick;

        public MyWindowCommandBar()
        {
            InitializeComponent();
            DataContext = this;
        }

        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);
            RefreshSettingBar();
            if (TitleBackground != null && MyGrid != null)
            {
                MyGrid.Background = TitleBackground;
            }
        }

        protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
        {
            base.OnPropertyChanged(change);
            if (change.Property == SettingBarProperty || change.Property == ShowIconProperty || change.Property == IconProperty)
            {
                RefreshSettingBar();
            }
            else if (change.Property == TitleBackgroundProperty && MyGrid != null)
            {
                MyGrid.Background = (IBrush?)change.NewValue;
            }
        }

        private void RefreshSettingBar()
        {
            if (MySettingBar == null)
            {
                return;
            }

            MySettingBar.Children.Clear();

            if (SettingBar != null)
            {
                MySettingBar.Children.Add(SettingBar);
            }
            else if (ShowIcon)
            {
                MySettingBar.Children.Add(new Image
                {
                    Source = Icon ?? LoadDefaultIcon(),
                    Stretch = Stretch.UniformToFill,
                    Margin = new Thickness(4),
                    Width = 28,
                    Height = 28
                });
            }
        }

        private static Bitmap LoadDefaultIcon() => new(AssetLoader.Open(new Uri("avares://MyAvalonia/Resources/Teradyne_T.ico")));

        private void HelpButton_Click(object? sender, RoutedEventArgs e) =>
            HelpClick?.Invoke(this, EventArgs.Empty);

        private void MinimizeButton_Click(object? sender, RoutedEventArgs e)
        {
            if (TopLevel.GetTopLevel(this) is Window window)
            {
                window.WindowState = WindowState.Minimized;
            }
        }

        private void MaximizeButton_Click(object? sender, RoutedEventArgs e)
        {
            if (TopLevel.GetTopLevel(this) is Window window)
            {
                window.WindowState = window.WindowState == WindowState.Maximized
                    ? WindowState.Normal
                    : WindowState.Maximized;
            }
        }

        private void CloseButton_Click(object? sender, RoutedEventArgs e)
        {
            if (TopLevel.GetTopLevel(this) is Window window)
            {
                window.Close();
            }
        }
    }
}
