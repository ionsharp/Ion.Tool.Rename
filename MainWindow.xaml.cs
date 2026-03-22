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

public class FileNameRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        if (string.IsNullOrEmpty(text))
            return new ValidationResult(false, "File name format isn't specified.");

        if (!text.Contains("{n}"))
            return new ValidationResult(false, "File name format must include {n}.");

        if (text.Any(c => Path.GetInvalidFileNameChars().Contains(c)))
            return new ValidationResult(false, "File name format includes invalid characters.");

        return ValidationResult.ValidResult;
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

        var result = int.TryParse(text, out int _);
        if (result)
            return ValidationResult.ValidResult;

        return new ValidationResult(false, "Input isn't a valid number.");
    }
}


public partial class MainWindow : Window, INotifyPropertyChanged
{
    /// <see cref="private"/>

    private readonly Lock Lock = new Lock();

    private const string TitleRename = "Rename";

    private const string TitleRenaming = "Renaming...";

    /// <see cref="public"/>

    public ObservableCollection<RenameOperation> Actions
    {
        get => field;
        private set
        {
            field = value;
            OnPropertyChanged();
        }
    } = new();


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
            Preview();
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


    public ObservableCollection<Exception> Log
    {
        get => field;
        private set
        {
            field = value;
            OnPropertyChanged();
        }
    } = new();

    public string LogTitle
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = LogTitleDefault;

    private const string LogTitleDefault = "0";


    public string PreviewTitle
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = PreviewTitleDefault;

    private const string PreviewTitleDefault = "0";


    public string NameFormat
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "{n}";

    public double NameFormatPadding
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 1;


    public bool IsBusy
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;

    public bool IsPreviewing
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;
    
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

    private IEnumerable<string> GetFilesSafe(string root, bool includeSubfolders)
    {
        if (!includeSubfolders)
        {
            return Directory.EnumerateFiles(root);
        }

        List<string> files = new();
        Stack<string> pending = new();
        pending.Push(root);

        while (pending.Count > 0)
        {
            var path = pending.Pop();
            try
            {
                // Add files in the current directory
                files.AddRange(Directory.EnumerateFiles(path));

                // Push subdirectories onto the stack to explore later
                foreach (var subdir in Directory.EnumerateDirectories(path))
                {
                    pending.Push(subdir);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Log the skipped folder so the user knows why it's missing
                // This is where you'd see 'System Volume Information' or 'Recovery'
                continue;
            }
            catch (DirectoryNotFoundException)
            {
                continue;
            }
        }
        return files;
    }

    private async void Preview()
    {
        lock (Lock)
        {
            if (IsBusy || PreviewAuto?.IsChecked != true)
                return;

            IsBusy = true;
            IsPreviewing = true;
        }

        if (string.IsNullOrWhiteSpace(Folder) || !Directory.Exists(Folder))
            goto Finish;

        var folder = Folder;
        var includeSub = IncludeSubfolders == true;

        // Use a local list for background processing
        var actions = new List<RenameOperation>();

        await Task.Run(() =>
        {
            try
            {
                var all = GetFilesSafe(folder, includeSub);
                var allGrouped = all.GroupBy(Path.GetDirectoryName);

                var at = int.TryParse(IncrementAt, out var x) ? x : 1;

                var iAt = at;
                var iBy = int.TryParse(IncrementBy, out var y) ? y : 1;
                
                var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                foreach (var group in allGrouped)
                {
                    var folderFiles = group.OrderBy(f => f).ToList();
                    foreach (var file in folderFiles)
                    {
                        ///Thread.Sleep(50);
                        string oldExt = Path.GetExtension(file);
                        string newExt = string.IsNullOrWhiteSpace(ExtensionOverride) ? oldExt : ExtensionOverride;

                        if (newExt.StartsWith('.'))
                            newExt = newExt[1..];

                        if (!string.IsNullOrEmpty(newExt))
                        {
                            newExt = ExtensionCasing switch
                            {
                                1 => newExt.ToLower(),
                                2 => newExt.ToUpper(),
                                3 => char.ToUpper(newExt[0]) + (newExt.Length > 1 ? newExt[1..].ToLower() : ""),
                                _ => newExt
                            };
                        }

                        newExt = "." + newExt;

                        int currentVal;
                        if (IncrementByExtension == true)
                        {
                            if (!counters.ContainsKey(oldExt))
                                counters[oldExt] = at;

                            currentVal = counters[oldExt];
                            counters[oldExt] += iBy;
                        }
                        else
                        {
                            currentVal = iAt;
                            iAt += iBy;
                        }

                        string incrementStr = currentVal.ToString().PadLeft(Convert.ToInt32(NameFormatPadding), '0');
                        string nameBase = NameFormat.Replace("{n}", incrementStr);

                        string newName = nameBase + newExt;
                        string newPath = Path.Combine(Path.GetDirectoryName(file)!, newName);

                        actions.Add(new RenameOperation
                        {
                            OldPath = file,
                            OldName = Path.GetFileName(file),
                            NewName = newName,
                            NewPath = newPath,
                            RelativeFolder = Path.GetRelativePath(folder, Path.GetDirectoryName(file)!)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => Log.Add(new Exception($"Preview Error: {ex.Message}")));
            }
        });

        Actions.Clear();
        foreach (var action in actions) 
            Actions.Add(action);

        PreviewTitle = $"{Actions.Count}";
        CountElement.Text = $"{Actions.Count} files";

        Finish:
        {
            lock (Lock)
            {
                IsPreviewing = false;
                IsBusy = false;
            }
        }
    }

    private async void Rename()
    {
        lock (Lock) 
        { 
            if (IsBusy || !Actions.Any()) 
                return;

            IsBusy = true;
            IsRenaming = true;
        }

        Title = TitleRenaming;

        Log.Clear();
        LogTitle = LogTitleDefault;

        if (Actions.Select(a => a.NewPath).Distinct(StringComparer.OrdinalIgnoreCase).Count() != Actions.Count)
        {
            lock (Lock) { IsRenaming = false; }

            var error = new Exception($"Current file name format causes collisions!");
            Log.Add(error);

            MessageBox.Show(error.Message, TitleRename, MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var result = MessageBox.Show($"Are you sure you want to rename {Actions.Count} files?", TitleRename, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (result != MessageBoxResult.Yes)
        {
            lock (Lock)
            {
                IsRenaming = false;
                IsBusy = false;
            }
            return;
        }

        List<Exception> errors = new();
        await Task.Run(() =>
        {
            try
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
                        File.Move(action.TempPath, action.NewPath);
                        Dispatcher.Invoke(() => action.Status = Brushes.Green);
                    }
                    catch (Exception f) 
                    { 
                        errors.Add(f);
                        Dispatcher.Invoke(() => action.Status = Brushes.Red);
                    }
                }
            }
            catch (Exception e) { errors.Add(e); }
        });

        Title = TitleRename;

        lock (Lock)
        {
            IsRenaming = false;
            IsBusy = false;
        }

        if (errors.Any())
        {
            foreach (var e in errors)
                Log.Add(e);

            LogTitle = $"{errors.Count}";
            MessageBox.Show($"One or more errors occurred while renaming. Check log for more details!", TitleRename);
        }
        else
        {
            MessageBox.Show("Rename finished successfully!", TitleRename);
        }

        Preview();
    }


    /// <see cref="event"/>

    private void AutoPreview_Updated(object sender, EventArgs e) => Preview();


    /// <see cref="RoutedEvent"/>

    private void OnBrowse(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "Select folder..." };
        if (dialog.ShowDialog() == true)
        {
            Folder = dialog.FolderName;
        }
    }

    private void OnPreview(object sender, RoutedEventArgs e) => Preview();

    private void OnRenamed(object sender, RoutedEventArgs e) => Rename();


    /// <see cref="INotifyPropertyChanged"/>

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class RenameOperation : INotifyPropertyChanged
{
    /// <see cref="public"/>

    public string OldPath { get; set; } = "";
    public string OldName { get; set; } = "";
    public string NewName
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "";
    public string NewPath
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "";
    public string TempPath { get; set; } = "";
    public string RelativeFolder { get; set; } = "";
    public Brush Status
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = Brushes.Orange;


    /// <see cref="INotifyPropertyChanged"/>

    public event PropertyChangedEventHandler? PropertyChanged;
    public void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}