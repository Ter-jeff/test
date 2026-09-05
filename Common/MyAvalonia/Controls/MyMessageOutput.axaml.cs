using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;

using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Styling;

namespace MyAvalonia.Controls
{

    [SuppressMessage("Maintainability", "CA1501:Avoid excessive inheritance", Justification = "Avalonia UI control inheritance depth")]
    public partial class MyMessageOutput : UserControl
    {
        public ObservableCollection<MessageItem> Messages { get; } = [];

        public MyMessageOutput()
        {
            InitializeComponent();
            DataContext = this;
        }

        public void Clear()
        {
            Messages.Clear();
        }

        public void AddMessage(int level, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            double baseFontSize = FontSize > 0 ? FontSize : 12;

            MessageItem item = level switch
            {
                //EnumMessageLevel.Warning
                1 => new MessageItem
                {
                    Text = "[Warning] " + message,
                    Foreground = ActualThemeVariant == ThemeVariant.Dark ? Brushes.Orange : Brushes.Orange,
                    FontSize = baseFontSize
                },
                //EnumMessageLevel.Error
                2 => new MessageItem
                {
                    Text = message,
                    Foreground = ActualThemeVariant == ThemeVariant.Dark ? Brushes.LightCoral : Brushes.Red,
                    FontSize = baseFontSize + 2
                },
                //EnumMessageLevel.Info
                3 => new MessageItem
                {
                    Text = message,
                    Foreground = ActualThemeVariant == ThemeVariant.Dark ? Brushes.LightGreen : Brushes.ForestGreen,
                    FontSize = baseFontSize
                },
                //EnumMessageLevel.StartPoint
                4 => new MessageItem
                {
                    Text = $"\n==================================\n{message}\n==================================",
                    Foreground = ActualThemeVariant == ThemeVariant.Dark ? Brushes.LightGreen : Brushes.ForestGreen,
                    FontSize = baseFontSize,
                    FontWeight = FontWeight.Bold
                },
                //EnumMessageLevel.EndPoint
                5 => new MessageItem
                {
                    Text = $"******* {message.PadLeft(((20 - message.Length) / 2) + message.Length),-20} *******",
                    Foreground = ActualThemeVariant == ThemeVariant.Dark ? Brushes.CornflowerBlue : Brushes.SteelBlue,
                    FontSize = baseFontSize + 4,
                    FontWeight = FontWeight.Bold,
                    FontFamily = new FontFamily("Courier New")
                },
                _ => new MessageItem
                {
                    Text = message,
                    Foreground = ActualThemeVariant == ThemeVariant.Dark ? Brushes.LightGray : Brushes.Black,
                    FontSize = baseFontSize
                }
            };

            Messages.Add(item);
            if (Messages.Count > 100)
            {
                Messages.RemoveAt(0);
            }

            Scroller.ScrollToEnd();
        }
    }

    public class MessageItem
    {
        public string Text { get; set; } = string.Empty;
        public IBrush Foreground { get; set; } = Brushes.Black;
        public double FontSize { get; set; } = 12;
        public FontWeight FontWeight { get; set; } = FontWeight.Normal;
        public FontFamily FontFamily { get; set; } = FontFamily.Default;
    }
}
