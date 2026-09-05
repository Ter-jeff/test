using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;

using MyAvalonia.Model;

namespace MyAvalonia.Controls
{
    [SuppressMessage("Maintainability", "CA1501:Avoid excessive inheritance", Justification = "Avalonia UI control inheritance depth")]
    public partial class HelpWindowFile : Window
    {
        private readonly Dictionary<string, string> _resources = [];
        private readonly ObservableCollection<object> _items = [];

        public HelpWindowFile()
        {
            InitializeComponent();
        }

        public HelpWindowFile SetDownload(string folder)
        {
            if (!Directory.Exists(folder))
            {
                return this;
            }

            string[] files = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
            var rows = new List<HelpFileRow>();

            foreach (string file in files)
            {
                string? dirPath = Path.GetDirectoryName(file);
                List<string> parts = dirPath?.Replace(folder, "")
                                    .Split(['\\', '/'], StringSplitOptions.RemoveEmptyEntries)
                                    .ToList() ?? [];
                string groupName = parts.Count > 0 ? string.Join("\\", parts) : string.Empty;
                string fileName = Path.GetFileName(file);
                string key = groupName + "\\" + fileName;

                rows.Add(new HelpFileRow { Key = key, FileName = fileName, GroupName = groupName, Select = true });
                _resources[key] = file;
            }

            // Manual group first, then the rest
            var sorted = rows.Where(r => r.GroupName.Equals("Manual", StringComparison.OrdinalIgnoreCase)).ToList();
            sorted.AddRange(rows.Where(r => !r.GroupName.Equals("Manual", StringComparison.OrdinalIgnoreCase)));

            string? currentGroup = null;
            foreach (HelpFileRow row in sorted)
            {
                if (row.GroupName.Contains("image", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (row.GroupName != currentGroup)
                {
                    currentGroup = row.GroupName;
                    _items.Add(new HelpGroupItem(string.IsNullOrEmpty(currentGroup) ? "Files" : currentGroup));
                }
                _items.Add(row);
            }

            FileList.ItemsSource = _items;
            return this;
        }

        private async void DownloadButton_Click(object? sender, RoutedEventArgs e)
        {
            IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Select Download Folder"
            });

            if (folders.Count == 0)
            {
                return;
            }

            string destFolder = folders[0].Path.LocalPath;

            foreach (HelpFileRow item in _items.OfType<HelpFileRow>())
            {
                if (!item.Select || !_resources.TryGetValue(item.Key, out string? sourcePath))
                {
                    continue;
                }

                string destFile = Path.Combine(destFolder, Path.GetFileName(item.Key));
                string? destDir = Path.GetDirectoryName(destFile);
                if (destDir != null && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                await CopyFileAsync(sourcePath, destFile);
            }

            Process.Start(new ProcessStartInfo("explorer.exe", destFolder) { UseShellExecute = true });
            Close();
        }

        private void Cancel_Click(object? sender, RoutedEventArgs e) => Close();

        private void CheckBox_PointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (!e.GetCurrentPoint(sender as Visual).Properties.IsRightButtonPressed)
            {
                return;
            }

            if (sender is not CheckBox { DataContext: HelpFileRow row })
            {
                return;
            }

            if (!_resources.TryGetValue(row.Key, out string? sourcePath))
            {
                return;
            }

            string destFile = Path.Combine(AppContext.BaseDirectory, sourcePath);
            OpenFile(destFile);
        }

        private static void OpenFile(string filePath)
        {
            try
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch
            {
                // No default app registered — open containing folder with file selected
                Process.Start(new ProcessStartInfo("explorer.exe")
                {
                    Arguments = $"/select,\"{filePath}\"",
                    UseShellExecute = true
                });
            }
        }

        private static async Task CopyFileAsync(string source, string dest)
        {
            using var input = new FileStream(source, FileMode.Open, FileAccess.Read);
            using FileStream output = File.OpenWrite(dest);
            await input.CopyToAsync(output);
        }
    }
}
