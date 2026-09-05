using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;

using MyAvalonia.Model;
using MyAvalonia.Utility;

namespace MyAvalonia.Controls
{
    [SuppressMessage("Maintainability", "CA1501:Avoid excessive inheritance", Justification = "Avalonia UI control inheritance depth")]
    public partial class MyTextBoxButton : UserControl
    {
        private EnumFileFilter _filter = EnumFileFilter.NoDefine;
        private bool _multiSelect;

        public static readonly StyledProperty<bool> MultiSelectProperty =
            AvaloniaProperty.Register<MyTextBoxButton, bool>(nameof(MultiSelect));

        public bool MultiSelect
        {
            get => GetValue(MultiSelectProperty);
            set => SetValue(MultiSelectProperty, value);
        }

        public static readonly StyledProperty<int> TextBlockWidthProperty =
            AvaloniaProperty.Register<MyTextBoxButton, int>(nameof(TextBlockWidth), 130);

        public int TextBlockWidth
        {
            get => GetValue(TextBlockWidthProperty);
            set => SetValue(TextBlockWidthProperty, value);
        }

        public static readonly StyledProperty<string?> HeaderProperty =
            AvaloniaProperty.Register<MyTextBoxButton, string?>(nameof(Header));

        public string? Header
        {
            get => GetValue(HeaderProperty);
            set => SetValue(HeaderProperty, value);
        }

        public static readonly StyledProperty<string?> TextProperty =
            AvaloniaProperty.Register<MyTextBoxButton, string?>(nameof(Text), defaultBindingMode: Avalonia.Data.BindingMode.TwoWay);

        public string? Text
        {
            get => GetValue(TextProperty);
            set => SetValue(TextProperty, value);
        }

        public event EventHandler? TextChanged;
        public event EventHandler? Click;

        public MyTextBoxButton()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void TextBoxButtonTextBox_TextChanged(object? sender, TextChangedEventArgs e)
        {
            TextChanged?.Invoke(this, EventArgs.Empty);
            Text = MyTextBox.Text;
            if (!string.IsNullOrEmpty(Text))
            {
                ValidatePaths();
            }
        }

        private void UpdateBorderBrush(bool exists)
        {
            MyBorder.BorderBrush = exists ? Brushes.SteelBlue : Brushes.Red;
            MyBorder.BorderThickness = exists ? new Thickness(1) : new Thickness(2);
        }

        private void ValidatePaths()
        {
            if (string.IsNullOrEmpty(Text))
            {
                return;
            }
            var arr = Text.Split(',').Select(s => s.Trim()).ToList();
            bool exists;

            if (_multiSelect || MultiSelect)
            {
                exists = _filter == EnumFileFilter.Folder
                    ? arr.All(Directory.Exists)
                    : _filter == EnumFileFilter.NoDefine
                        ? arr.All(p => File.Exists(p) || Directory.Exists(p))
                        : arr.All(File.Exists);
            }
            else
            {
                exists = _filter == EnumFileFilter.Folder
                    ? Directory.Exists(Text)
                    : _filter == EnumFileFilter.NoDefine
                        ? File.Exists(Text) || Directory.Exists(Text)
                        : File.Exists(Text);
            }

            UpdateBorderBrush(exists);
        }

        private void Button_Click(object? sender, RoutedEventArgs e)
        {
            Click?.Invoke(this, EventArgs.Empty);
        }

        public async Task<bool> FileSelectAsync(MyWindow window, EnumFileFilter filter, Action<string, int, int, string> report, bool multiSelect = false)
        {
            _filter = filter;
            _multiSelect = multiSelect;

            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null)
            {
                return false;
            }

            List<string> patterns = GetPatterns(filter);
            var fileTypeFilter = new List<FilePickerFileType>
            {
                new(FileFilter.GetFilter(filter).Split('|')[0]) { Patterns = patterns }
            };

            string? initialDirectory = GetInitialDirectory(multiSelect);
            IStorageFolder? startFolder = initialDirectory != null
                ? await topLevel.StorageProvider.TryGetFolderFromPathAsync(initialDirectory)
                : null;

            IReadOnlyList<IStorageFile> files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Select File",
                AllowMultiple = multiSelect,
                FileTypeFilter = fileTypeFilter,
                SuggestedStartLocation = startFolder
            });

            if (files.Count == 0)
            {
                Text = string.Empty;
                return false;
            }

            if (multiSelect)
            {
                var paths = files.Select(f => f.Path.LocalPath).ToList();
                foreach (string p in paths)
                {
                    report($"{Header?.Split('\n').FirstOrDefault()} : {p}", 0, 0, "");
                }
                Text = string.Join(",", paths);
                window.DefaultPath = Path.GetDirectoryName(paths.Last());
            }
            else
            {
                Text = files[0].Path.LocalPath;
                report($"{Header?.Split('\n').FirstOrDefault()} : {Text}", 0, 0, "");
                window.DefaultPath = Path.GetDirectoryName(Text);
            }

            return true;
        }

        public async Task<bool> FolderSelectAsync(MyWindow window, Action<string, int, int, string> report, bool multiSelect = false)
        {
            _filter = EnumFileFilter.Folder;
            _multiSelect = multiSelect;

            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null)
            {
                return false;
            }

            string? initialDirectory = GetInitialDirectory(multiSelect);
            IStorageFolder? startFolder = initialDirectory != null
                ? await topLevel.StorageProvider.TryGetFolderFromPathAsync(initialDirectory)
                : null;

            IReadOnlyList<IStorageFolder> folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Select Folder",
                AllowMultiple = multiSelect,
                SuggestedStartLocation = startFolder
            });

            if (folders.Count == 0)
            {
                Text = string.Empty;
                return false;
            }

            if (multiSelect)
            {
                var paths = folders.Select(f => f.Path.LocalPath).ToList();
                foreach (string p in paths)
                {
                    report($"{Header?.Split('\n').FirstOrDefault()} : {p}", 0, 0, "");
                }
                Text = string.Join(",", paths);
                window.DefaultPath = Text;
            }
            else
            {
                Text = folders[0].Path.LocalPath;
                report($"{Header?.Split('\n').FirstOrDefault()} : {Text}", 0, 0, "");
                window.DefaultPath = Text;
            }

            return true;
        }

        private string? GetInitialDirectory(bool multiSelect)
        {
            if (string.IsNullOrEmpty(Text))
            {
                return null;
            }
            string first = multiSelect ? Text.Split(',').First() : Text;
            return Path.GetDirectoryName(first);
        }

        private static List<string> GetPatterns(EnumFileFilter filter) => filter switch
        {
            EnumFileFilter.Igxl => ["*.igxl"],
            EnumFileFilter.Excel => ["*.xlsx", "*.xlsm"],
            EnumFileFilter.ExcelOld => ["*.xlsx", "*.xlsm", "*.xls"],
            EnumFileFilter.Txt => ["*.txt"],
            EnumFileFilter.Csv => ["*.csv"],
            EnumFileFilter.XmlFile => ["*.xml"],
            EnumFileFilter.YamlFile => ["*.yaml"],
            _ => ["*.*"]
        };
    }
}
