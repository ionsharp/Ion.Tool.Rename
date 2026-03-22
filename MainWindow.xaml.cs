using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace Ion.Tool.Rename;

public class InverseBooleanConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool i ? !i : false;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool i ? !i : false;
}


public class FileRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        if (string.IsNullOrEmpty(text))
            return ValidationResult.ValidResult;

        return new ValidationResult(false, "Changing a file extension can cause a file to become unreadable.");
    }
}

public class FolderRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        if (Directory.Exists(text))
            return ValidationResult.ValidResult;

        return new ValidationResult(false, "Folder doesn't exist!");
    }
}

public class NumberRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        var result = double.TryParse(text, out double _);
        if (result)
            return ValidationResult.ValidResult;

        return new ValidationResult(false, "Input isn't a valid number.");
    }
}


public partial class MainWindow : Window, INotifyPropertyChanged
{
    /// <see cref="private"/>

    public ObservableCollection<RenameOperation> Actions
    {
        get => field;
        private set
        {
            field = value;
            OnPropertyChanged();
        }
    } = new();

    /// <see cref="public"/>

    public int ExtensionCasing
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 1;

    public string ExtensionOverride
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "";

    public string Folder
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "";
    
    public bool? IncludeSubfolders
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;

    public string IncrementAt
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "1";

    public string IncrementBy
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "1";

    public bool? IncrementByExtension
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;

    public double NameFormatPadding
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 1;
    

    private readonly Lock Lock = new Lock();

    public bool IsRenaming
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;


    /// <see cref="MainWindow"/>

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
    }


    /// <see cref="private"/>

    private void Preview()
    {
        if (string.IsNullOrWhiteSpace(Folder) || !Directory.Exists(Folder)) return;
        Actions.Clear();

        var folder = Folder;
        var folderOption = IncludeSubfolders == true ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

        var allFiles = Directory.GetFiles(folder, "*.*", folderOption).GroupBy(Path.GetDirectoryName);

        foreach (var group in allFiles)
        {
            var folderFiles = group.OrderBy(f => f).ToList();
            
            var iAt = int.TryParse(IncrementAt, out var x) ? x : 1; /// Global counter
            var iBy = int.TryParse(IncrementBy, out var y) ? y : 1;

            /// Counters per extension
            var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var file in folderFiles)
            {
                string oldExt = Path.GetExtension(file);
                string newExt = string.IsNullOrWhiteSpace(ExtensionOverride) ? oldExt : ExtensionOverride;

                if (newExt.StartsWith('.')) 
                    newExt = newExt[1..];

                newExt = ExtensionCasing switch
                {
                    1 => newExt.ToLower(),
                    2 => newExt.ToUpper(),
                    3 => char.ToUpper(newExt[0]) + newExt[1..].ToLower(),
                    _ => newExt
                };

                newExt = "." + newExt;

                // Determine Increment
                int currentVal;
                if (IncrementByExtension == true)
                {
                    if (!counters.ContainsKey(oldExt)) 
                        counters[oldExt] = iAt;

                    currentVal = counters[oldExt];
                    counters[oldExt] += iBy;
                }
                else
                {
                    currentVal = iAt;
                    iAt += iBy;
                }

                string incrementStr = currentVal.ToString().PadLeft(Convert.ToInt32(NameFormatPadding), '0');
                string nameBase = PatternBox.Text.Replace("{n}", incrementStr);

                string newName = nameBase + newExt;
                string newPath = Path.Combine(Path.GetDirectoryName(file)!, newName);

                Actions.Add(new RenameOperation
                {
                    OldPath = file,
                    OldName = Path.GetFileName(file),
                    NewName = newName,
                    NewPath = newPath,
                    RelativeFolder = Path.GetRelativePath(folder, Path.GetDirectoryName(file)!)
                });
            }
        }
        CountElement.Text = $"{Actions.Count} files";
    }

    private async void Rename()
    {
        lock (Lock) { if (IsRenaming) return; }
        IsRenaming = true;

        if (!Actions.Any()) return;

        var result = MessageBox.Show($"Are you sure you want to rename {Actions.Count} files?",
            "Rename", MessageBoxButton.YesNo, MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            lock (Lock) { IsRenaming = false; }
            return;
        }

        await Task.Run(() =>
        {
            /// 1. Rename temporarily to avoid collisions
            foreach (var action in Actions)
            {
                string temp = action.OldPath + ".tmp_" + Guid.NewGuid().ToString("N");

                File.Move(action.OldPath, temp);
                action.TempPath = temp;
            }

            /// 2. Final rename
            foreach (var action in Actions)
            {
                try
                {
                    if (File.Exists(action.NewPath))
                        File.Delete(action.NewPath); // Final safety

                    File.Move(action.TempPath, action.NewPath);
                }
                catch (Exception e)
                {
                    Dispatcher.Invoke(() => MessageBox.Show($"Couldn't rename {action.OldPath}: {e.Message}"));
                }
            }
        });

        lock (Lock) { IsRenaming = false; }
        MessageBox.Show("Rename finished!", "Rename");

        Preview();
    }


    /// <see cref="event"/>

    private void AutoPreview_Updated(object sender, EventArgs e)
    {
        if (PreviewAuto?.IsChecked == true)
            Preview();
    }


    /// <see cref="RoutedEvent"/>

    private void OnBrowse(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "Select folder..." };
        if (dialog.ShowDialog() == true)
        {
            Folder = dialog.FolderName;
            Preview();
        }
    }

    private void OnPreview(object sender, RoutedEventArgs e) => Preview();

    private void OnRenamed(object sender, RoutedEventArgs e) => Rename();


    /// <see cref="INotifyPropertyChanged"/>

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class RenameOperation
{
    public string OldPath { get; set; } = "";
    public string OldName { get; set; } = "";
    public string NewName { get; set; } = "";
    public string NewPath { get; set; } = "";
    public string TempPath { get; set; } = "";
    public string RelativeFolder { get; set; } = "";
    public Brush StatusColor => File.Exists(NewPath) ? Brushes.Orange : Brushes.Green;
}