using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Ion.Tool.Rename;

public record class LetterNumberStyle
{
}

/// <see cref="MainWindow"/>
#region

public partial class MainWindow : Window, INotifyPropertyChanged
{
    /// <see cref="private"/>

    private bool Dirty = false;

    private readonly Lock Lock = new Lock();

    private const string TitleRename = "Rename";

    private const string TitleRenaming = "Renaming...";


    /// <see cref="public"/>

    public ObservableCollection<RenameTask> Actions
    {
        get => field;
        private set
        {
            field = value;
            OnPropertyChanged();
        }
    } = new();


    /// <remarks>In milliseconds!</remarks>
    public string DelayRename
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "0"; /// 1000
    
    /// <remarks>In milliseconds!</remarks>
    public string DelayUpdate
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "0"; /// 1000


    public int FileCount
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 0;

    public string FileCountText
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "0 files";


    public int FileExtensionCasing
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 1;

    public bool? FileExtensionOverride
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;

    public string FileExtensionOverrideNew
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "";


    public string FileNameIncrementAt
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "1";

    public string FileNameIncrementBy
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "1";

    public bool? FileNameIncrementByExtension
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;


    public string FileNameFormat
    {
        get => field;
        set
        {
            field = System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " ");
            OnPropertyChanged();

            Update();
        }
    } = FileNameFormatSymbol;

    public const string FileNameFormatSymbol = "{n}";


    public int FileNameNumber
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 1;

    public double FileNameNumberPadding
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 1;

    public string FileNameNumberPaddingCharacter
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "0";


    public string Folder
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();

            if (Directory.Exists(Folder))
                FolderPreview = Thumbnail.GetLarge(Folder);

            _ = Update();
        }
    } = "";

    public ImageSource FolderPreview
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    }

    public bool? FolderSub
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
            field?.CollectionChanged -= OnLogChanged;

            field = value;
            OnPropertyChanged();

            LogCount = field?.Count ?? 0;
            field?.CollectionChanged += OnLogChanged;
        }
    } = new();

    public int LogCount
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = 0;

    public bool? LogEnable
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;


    public bool IsBusy
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;

    public bool IsUpdating
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

    private void LogError(Exception e)
    {
        Dispatcher.Invoke(() =>
        {
            if (LogEnable is not true) return;
            Log.Add(e);
        });
    }

    private void LogErrors(IEnumerable<Exception> e)
    {
        Dispatcher.Invoke(() =>
        {
            if (LogEnable is not true) return;
            Log = new ObservableCollection<Exception>(e);
        });
    }


    private static string ToLetter(int n)
    {
        const int columns = 26;
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        //ceil(log26(Int32.Max))
        const int digitMaximum = 7;

        n = n < 0 ? 0 : n;
        if (n <= columns)
            return letters[n - 1].ToString();

        var result = new StringBuilder().Append(' ', digitMaximum);

        var current = n;
        var offset = digitMaximum;
        while (current > 0)
        {
            result[--offset] = letters[--current % columns];
            current /= columns;
        }

        return result.ToString(offset, digitMaximum - offset);
    }

    private static string ToNumber(int n, char[] characters)
    {
        string result = string.Empty;
        int targetBase = characters.Length;

        do
        {
            result = characters[n % targetBase] + result;
            n = n / targetBase;
        }
        while (n > 0);

        return result;
    }


    private string GetFileExtension(string file)
    {
        string oldExt = Path.GetExtension(file);
        string newExt = FileExtensionOverride is not true || string.IsNullOrWhiteSpace(FileExtensionOverrideNew) ? oldExt : FileExtensionOverrideNew;

        if (newExt.StartsWith('.')) newExt = newExt[1..];
        if (!string.IsNullOrEmpty(newExt))
        {
            newExt = FileExtensionCasing switch
            {
                1 => newExt.ToLower(),
                2 => newExt.ToUpper(),
                3 => char.ToUpper(newExt[0]) + (newExt.Length > 1 ? newExt[1..].ToLower() : ""),
                _ => newExt
            };
            newExt = "." + newExt;
        }
        return newExt;
    }

    private string GetFileName(int number)
    {
        var padWith = FileNameNumberPaddingCharacter?.Length > 0 ? FileNameNumberPaddingCharacter[0] : '0';

        var oldNumber = $"{number}";
        var newNumber = oldNumber;
        switch (FileNameNumber)
        {
            /// (02) Binary
            case 0:
                return ToNumber(number, new char[] { '0', '1' });
            /// (10) Decimal
            case 1:
                newNumber = oldNumber;
                break;
            /// (16) Hexadecimal
            case 2:
                newNumber = ToNumber(number, new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F' });
                break;
            /// Letter
            case 3:
            case 4:
                newNumber = ToLetter(number);
                newNumber = FileNameNumber == 3 ? newNumber.ToLower() : newNumber.ToUpper();
                break; 
        }

        var result = newNumber.PadLeft(Convert.ToInt32(FileNameNumberPadding), padWith);
        return FileNameFormat.Replace(FileNameFormatSymbol, result).Trim();
    }

    private IEnumerable<string> GetFiles(string root, bool includeSub)
    {
        if (!includeSub)
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
                /// Add files in the current directory
                files.AddRange(Directory.EnumerateFiles(path));

                /// Push subdirectories onto the stack to explore later
                foreach (var subdir in Directory.EnumerateDirectories(path))
                {
                    pending.Push(subdir);
                }
            }
            catch (UnauthorizedAccessException)
            {
                /// Log the skipped folder so the user knows why it's missing
                /// This is where you'd see 'System Volume Information' or 'Recovery'
                continue;
            }
            catch (DirectoryNotFoundException)
            {
                continue;
            }
        }
        return files;
    }


    private async Task Rename()
    {
        lock (Lock)
        {
            if (IsBusy || !Actions.Any())
                return;

            IsBusy = true;
            IsRenaming = true;
        }

        Title = TitleRenaming;

        /// Internal collisions (new names hitting each other)
        if (Actions.Select(a => a.NewPath).Distinct(StringComparer.OrdinalIgnoreCase).Count() != Actions.Count)
        {
            RenameAbort($"Current file name format causes collisions!");
            return;
        }

        /// External collisions (new names hitting files already on disk that aren't part of this set)
        var currentFiles = new HashSet<string>(Actions.Select(a => a.OldPath), StringComparer.OrdinalIgnoreCase);
        foreach (var action in Actions)
        {
            /// If the destination exists AND it's not one of the files we are currently renaming
            if (File.Exists(action.NewPath) && !currentFiles.Contains(action.NewPath))
            {
                RenameAbort($"Collision detected: The file '{Path.GetFileName(action.NewPath)}' already exists in the destination.");
                return;
            }
        }

        var result = MessageBox.Show($"Are you sure you want to rename {Actions.Count} files?", TitleRename, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (result != MessageBoxResult.Yes)
        {
            RenameReset();
            return;
        }

        List<Exception> errors = new();
        bool criticalFailure = false;

        await Task.Run(() =>
        {
            int.TryParse(DelayRename, out int delay);
            Thread.Sleep(delay);

            // The undo stack remains our safety net if a move fails halfway through
            Stack<(string Current, string Original)> undoStack = new();
            try
            {
                int tempCounter = 1;

                /// 1. Rename temporarily to avoid collisions
                foreach (var action in Actions)
                {
                    try
                    {
                        string directory = Path.GetDirectoryName(action.OldPath)!;

                        string tempName = $"{tempCounter++}.tmp";
                        string tempPath = Path.Combine(directory, tempName);

                        // If by some freak accident 1.tmp already exists, we skip ahead
                        while (File.Exists(tempPath))
                        {
                            tempPath = Path.Combine(directory, $"{tempCounter++}.tmp");
                        }

                        File.Move(action.OldPath, tempPath);

                        action.TempPath = tempPath;
                        undoStack.Push((tempPath, action.OldPath));
                    }
                    catch (Exception ex)
                    {
                        errors.Add(new Exception($"Critical failure moving {action.OldName} to temp: {ex.Message}"));
                        criticalFailure = true;
                        break;
                    }
                }

                /// 2. Final rename
                if (!criticalFailure)
                {
                    foreach (var action in Actions)
                    {
                        try
                        {
                            File.Move(action.TempPath, action.NewPath);

                            /// Update the stack: the file is now at NewPath, needs to go back to OldPath on failure
                            /// We pop the old (temp->old) and push the new (new->old)
                            undoStack.Pop();
                            undoStack.Push((action.NewPath, action.OldPath));

                            Dispatcher.Invoke(() => action.Status = Brushes.Green);
                        }
                        catch (Exception f)
                        {
                            errors.Add(new Exception($"Final move failed for {action.NewName}: {f.Message}"));
                            Dispatcher.Invoke(() => action.Status = Brushes.Red);
                            criticalFailure = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception e) { errors.Add(e); criticalFailure = true; }

            /// ROLLBACK: If any critical failure occurred, undo everything in the stack
            if (criticalFailure)
            {
                while (undoStack.Count > 0)
                {
                    var operation = undoStack.Pop();
                    try
                    {
                        if (File.Exists(operation.Current))
                            File.Move(operation.Current, operation.Original);
                    }
                    catch { /* Final effort failed - logged if necessary */ }
                }
            }
        });

        Title = TitleRename;
        RenameReset();

        if (errors.Any())
        {
            LogErrors(errors);

            var message = $"One or more errors occurred while renaming.";
            message += LogEnable is true ? " Check log for more details!" : "";

            MessageBox.Show(message, TitleRename);
        }
        else
        {
            MessageBox.Show("Rename finished successfully!", TitleRename);
        }

        Update();
    }

    private void RenameAbort(string message)
    {
        RenameReset();

        var error = new Exception(message);
        LogError(error);

        MessageBox.Show(error.Message, TitleRename, MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void RenameReset()
    {
        lock (Lock) { IsRenaming = false; IsBusy = false; }
    }


    private async Task Update()
    {
        lock (Lock)
        {
            if (IsBusy)
            {
                Dirty = true;
                return;
            }

            IsBusy = true;
            IsUpdating = true;
            Dirty = false;
        }

        Log.Clear();

        try
        {
            if (string.IsNullOrWhiteSpace(Folder))
            {
                LogError(new ArgumentException($"Folder wasn't specified."));
                return;
            }

            if (!Directory.Exists(Folder))
            {
                LogError(new DirectoryNotFoundException($"Folder '{Folder}' does not exist."));
                return;
            }

            var a = Folder;
            var b = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            var relativePath = Path.GetRelativePath(b, a);

            /// Ignore any folder that is not a subfolder of the current user folder. For safety, user folder itself is ignored!
            bool isOutside = relativePath.StartsWith("..") || Path.IsPathRooted(relativePath);

            if (a == b || isOutside)
            {
                LogError(new AccessViolationException("Folder is not a subfolder of the current user folder."));
                return;
            }

            var folder = Folder;
            var includeSub = FolderSub == true;
            var actions = new List<RenameTask>();

            await Task.Run(() =>
            {
                int.TryParse(DelayUpdate, out int delay);
                Thread.Sleep(delay);
                try
                {
                    var all = GetFiles(folder, includeSub);
                    var allGrouped = all.GroupBy(Path.GetDirectoryName);

                    var at = int.TryParse(FileNameIncrementAt, out var x) ? x : 1;

                    var iAt = at;
                    var iBy = int.TryParse(FileNameIncrementBy, out var y) ? y : 1;

                    var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (var group in allGrouped)
                    {
                        var folderFiles = group.OrderBy(f => f).ToList();
                        foreach (var file in folderFiles)
                        {
                            string oldExt = Path.GetExtension(file);
                            string newExt = GetFileExtension(file);

                            int currentVal;
                            if (FileNameIncrementByExtension == true)
                            {
                                if (!counters.ContainsKey(oldExt)) counters[oldExt] = at;
                                currentVal = counters[oldExt];
                                counters[oldExt] += iBy;
                            }
                            else
                            {
                                currentVal = iAt;
                                iAt += iBy;
                            }

                            string newName = GetFileName(currentVal) + newExt;
                            string newPath = Path.Combine(Path.GetDirectoryName(file)!, newName);

                            actions.Add(new RenameTask
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
                    LogError(new Exception($"Update Error: {ex.Message}"));
                }
            });

            Actions = new ObservableCollection<RenameTask>(actions);

            FileCount = Actions.Count;
            FileCountText = FileCount == 1 ? "1 file" : $"{FileCount} files";
        }
        finally
        {
            /// --- CLEANUP & RE-RUN CHECK ---
            bool runAgain = false;
            lock (Lock)
            {
                IsUpdating = false;
                IsBusy = false;

                /// If someone tried to Update while we were working, 
                /// we capture that intent here.
                if (Dirty)
                {
                    runAgain = true;
                    Dirty = false;
                }
            }

            /// If a refresh was requested, call Update again.
            /// Using Task.Yield() or a small delay prevents a massive stack of calls.
            if (runAgain)
            {
                Dispatcher.BeginInvoke(new Action(() => Update()));
            }
        }
    }


    /// <see cref="event"/>

    private void OnLogChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        LogCount = Log.Count;
    }

    private void OnUpdated(object sender, EventArgs e) => _ = Update();


    /// <see cref="RoutedEvent"/>

    private void OnBrowse(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "Select folder..." };
        if (dialog.ShowDialog() == true)
            Folder = dialog.FolderName;
    }

    private void OnUpdate(object sender, RoutedEventArgs e) => _ = Update();

    private void OnRenamed(object sender, RoutedEventArgs e) => _ = Rename();


    /// <see cref="INotifyPropertyChanged"/>

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

#endregion

/// <see cref="INotifyPropertyChanged"/>
#region

public class RenameTask : INotifyPropertyChanged
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

#endregion

/// <see cref="IValueConverter"/>
#region

public class InverseBooleanConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool i ? !i : false;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool i ? !i : false;
}

public class ThumbnailConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is string i ? Thumbnail.GetLarge(i) : Binding.DoNothing;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => Binding.DoNothing;
}

public class ToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is object i ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => Binding.DoNothing;
}

#endregion

/// <see cref="ValidationRule"/>
#region

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

public class FileNameFormatRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        if (string.IsNullOrEmpty(text))
            return new ValidationResult(false, "File name format isn't specified.");

        if (!text.Contains(MainWindow.FileNameFormatSymbol))
            return new ValidationResult(false, $"File name format must include the symbol '{MainWindow.FileNameFormatSymbol}'.");

        text = text.Replace(MainWindow.FileNameFormatSymbol, "");
        if (text.Any(c => Path.GetInvalidFileNameChars().Contains(c)))
            return new ValidationResult(false, "File name format includes invalid characters.");

        return ValidationResult.ValidResult;
    }
}

public class FileNameRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;
        if (text?.Any(c => Path.GetInvalidFileNameChars().Contains(c)) == true)
            return new ValidationResult(false, "Invalid characters.");

        return ValidationResult.ValidResult;
    }
}

public class MinimumRule : ValidationRule
{
    public int Minimum { get; set; }

    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;
        if (text?.Length < Minimum)
            return new ValidationResult(false, "Not enough characters.");

        return ValidationResult.ValidResult;
    }
}

public class FolderRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;
        if (string.IsNullOrEmpty(text) || Directory.Exists(text))
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

        return new ValidationResult(false, "Not a valid number.");
    }
}

#endregion