using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.VisualTree;

using MyAvalonia.Model;

namespace MyAvalonia.Controls
{
    [SuppressMessage("Maintainability", "CA1501:Avoid excessive inheritance", Justification = "Avalonia UI control inheritance depth")]
    public abstract class MyWindow : Window
    {
        public static readonly StyledProperty<bool> ShowHelpProperty = AvaloniaProperty.Register<MyWindow, bool>(nameof(ShowHelp), true);

        public static readonly StyledProperty<bool> ShowIconProperty = AvaloniaProperty.Register<MyWindow, bool>(nameof(ShowIcon), true);

        public static readonly StyledProperty<IBrush?> TitleBackgroundProperty = AvaloniaProperty.Register<MyWindow, IBrush?>(nameof(TitleBackground));

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

        public IBrush? TitleBackground
        {
            get => GetValue(TitleBackgroundProperty);
            set => SetValue(TitleBackgroundProperty, value);
        }

        private string? _defaultPath;

        public string? DefaultPath
        {
            get => _defaultPath;
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _defaultPath = value;
                }
            }
        }

        public MySettings MySettings { get; set; } = new();

        protected override Type StyleKeyOverride => typeof(MyWindow);

        public event EventHandler<EventArgs>? HelpButtonClick;

        protected abstract void CheckStatus();

        protected void RaiseHelpButtonClick() => HelpButtonClick?.Invoke(this, EventArgs.Empty);

        protected async void HelpButton_Clicked(string folder)
        {
            var helpWindow = new HelpWindowFile();
            helpWindow.SetDownload(folder);
            await helpWindow.ShowDialog(this);
        }

        protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
        {
            base.OnApplyTemplate(e);

            MyWindowCommandBar? commandBar = e.NameScope.Find<MyWindowCommandBar>("MyWindowCommandBar");
            if (commandBar == null)
            {
                return;
            }

            commandBar.HelpClick += (_, _) => RaiseHelpButtonClick();
            commandBar.DoubleTapped += (_, _) => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
            commandBar.PointerPressed += OnCommandBarPointerPressed;
        }

        private void OnCommandBarPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (!e.GetCurrentPoint(sender as Visual).Properties.IsLeftButtonPressed)
            {
                return;
            }

            if (e.Source is Visual source && (source is Button || source.GetVisualAncestors().OfType<Button>().Any()))
            {
                return;
            }

            BeginMoveDrag(e);
        }

        protected static string GetVersion(Assembly assembly)
        {
            string fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version
                ?? assembly.GetName().Version?.ToString()
                ?? "0.0.0.0";
            return " " + fileVersion;
        }

        public void WriteSettings(Control root, string filePath)
        {
            var mySetting = new MySetting { ClassType = root.GetType().ToString() };
            foreach (Control control in FindAllControls(root))
            {
                string controlType = control.GetType().ToString();
                string controlName = control.Name ?? string.Empty;
                if (string.IsNullOrEmpty(controlName))
                {
                    continue;
                }

                string? content = null;
                if (control is TextBox textBox && !controlName.Equals("MyTextBox", StringComparison.OrdinalIgnoreCase))
                {
                    content = textBox.Text;
                }
                else if (control is MyTextBoxButton myTextBoxButton)
                {
                    content = myTextBoxButton.Text;
                }
                else if (control is ComboBox comboBox)
                {
                    content = comboBox.SelectedIndex.ToString();
                }
                else if (control is RadioButton radioButton)
                {
                    content = radioButton.IsChecked.ToString();
                }
                else if (control is CheckBox checkBox)
                {
                    content = checkBox.IsChecked.ToString();
                }

                if (content != null)
                {
                    AddControl(mySetting, controlType, controlName, content);
                }
            }
            (MySettings ??= new MySettings()).Write(filePath, mySetting);
        }

        public void LoadSettings(string fileName)
        {
            try
            {
                if (!File.Exists(fileName))
                {
                    return;
                }

                MySettings = Json.Read<MySettings>(fileName) ?? new MySettings();
                if (!MySettings.Collection.Exists(x => x.ClassType == ToString()))
                {
                    return;
                }

                var allSetting = (MySetting)MySettings.Collection.Find(x => x.ClassType == ToString())!;
                foreach (MyControl myControl in allSetting.MyControls)
                {
                    string controlName = myControl.ControlName;
                    string content = myControl.Content;

                    Control? control = this.FindControl<Control>(controlName);
                    if (control == null)
                    {
                        continue;
                    }

                    if (control is TextBox textBox)
                    {
                        textBox.Text = content;
                    }
                    else if (control is MyTextBoxButton myTextBoxButton)
                    {
                        myTextBoxButton.Text = content;
                    }
                    else if (control is ComboBox comboBox && int.TryParse(content, out int idx))
                    {
                        comboBox.SelectedIndex = idx;
                    }
                    else if (control is RadioButton radioButton && bool.TryParse(content, out bool rb))
                    {
                        radioButton.IsChecked = rb;
                    }
                    else if (control is CheckBox checkBox && bool.TryParse(content, out bool cb))
                    {
                        checkBox.IsChecked = cb;
                    }
                }
            }
            catch (Exception)
            {
                // skip corrupt settings
            }
            finally
            {
                CheckStatus();
            }
        }

        private static void AddControl(MySetting mySetting, string controlType, string controlName, string content)
        {
            if (!mySetting.MyControls.Exists(x => x.Content == content && x.ControlName == controlName && x.ControlType == controlType))
            {
                mySetting.MyControls.Add(new MyControl { ControlType = controlType, ControlName = controlName, Content = content });
            }
        }

        private static IEnumerable<Control> FindAllControls(Visual parent)
        {
            foreach (Visual child in parent.GetVisualChildren())
            {
                if (child is Control control)
                {
                    yield return control;
                }
                if (child is Visual visual)
                {
                    foreach (Control descendant in FindAllControls(visual))
                    {
                        yield return descendant;
                    }
                }
            }
        }
    }
}
