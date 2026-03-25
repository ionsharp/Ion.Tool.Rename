using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace Ion.Tool.Rename;

/// <see cref="MainWindow"/>
#region

public partial class MainWindow : Window, INotifyPropertyChanged
{
    /// <see cref="MemberTypes.Property"/>
    #region

    private bool Dirty = false;

    private readonly Lock Lock = new Lock();

    private const string TitleRename = "Rename";
    private const string TitleRenaming = "Renaming...";


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
    } = true;


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


    public ObservableCollection<Operation> Files
    {
        get => field;
        private set
        {
            field = value;
            OnPropertyChanged();
        }
    } = [];

    public ObservableCollection<FilterItem> Filters { get; } = [];


    public string Folder
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();

            if (Directory.Exists(Folder))
                FolderPreview = Thumbnail.GetIcon(Folder, ThumbnailSize.Jumbo);

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


    public ObservableCollection<Message> Log
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
    }

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
    } = true;


    public bool TestAutoSelect
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = true;
    public string TestFileCount
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = "50";
    public string TestFileExtension
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = ".txt";
    public string TestFolderName
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = "Test";
    public string TestNameLength
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = "12";
    public string TestTargetRoot
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


    public ObservableCollection<OperationInfo> RedoActions { get; private set; } = new();
    public ObservableCollection<OperationInfo> UndoActions { get; private set; } = new();


    public bool IsDraggingOver
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = false;
    public string IsDraggingOverFolder
    {
        get => field;
        set { field = value; OnPropertyChanged(); }
    } = "";


    public bool IsBusy
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;
    public bool IsRedoing
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
    public bool IsTesting
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = false;
    public bool IsUndoing
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

    #endregion

    /// <see cref="MainWindow"/>
    #region

    public MainWindow()
    {
        PresetManager.Load();

        Filters.CollectionChanged += (s, e) => Update();
        Log = new ObservableCollection<Message>();

        DataContext = this;
        InitializeComponent();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        PresetManager.Save();
        base.OnClosing(e);
    }

    #endregion

    /// <see cref="MemberTypes.Method"/>
    #region

    private void LogMessage(Exception e)
    {
        Dispatcher.Invoke(() =>
        {
            if (LogEnable is not true) return;
            Log.Add(new Message(e));
        });
    }

    private void LogMessages(IEnumerable<Exception> e)
    {
        Dispatcher.Invoke(() =>
        {
            if (LogEnable is not true) return;
            Log = new ObservableCollection<Message>(e.Select(i => new Message(i)));
        });
    }


    private IEnumerable<string> Filter(IEnumerable<string> files) => files.Where(i => Filters.All(filter => filter.IsMatch(new FileInfo(i))));


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
            return Directory.EnumerateFiles(root);

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
                    pending.Push(subdir);
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


    private bool IsFolderChildOf(string child, string parent)
    {
        var pathRelative = Path.GetRelativePath(parent, child);

        if (child == parent || pathRelative.StartsWith("..") || Path.IsPathRooted(pathRelative))
        {
            LogMessage(new Exception(Messages.FolderNotValid));
            return false;
        }
        return true;
    }

    private bool IsFolderSafe(string path)
    {
        /// Windows MAX_PATH is 260. We use 255 for a safety buffer.
        if (path.Length >= 255)
            return false;

        return true;
    }


    private async Task Rename()
    {
        lock (Lock)
        {
            if (IsBusy || !Files.Any())
                return;

            IsBusy = true;
            IsRenaming = true;
        }

        Title = TitleRenaming;

        /// Internal collisions (new names hitting each other)
        if (Files.Select(a => a.NewPath).Distinct(StringComparer.OrdinalIgnoreCase).Count() != Files.Count)
        {
            RenameAbort($"Current file name format causes collisions!");
            return;
        }

        /// External collisions (new names hitting files already on disk that aren't part of this set)
        var currentFiles = new HashSet<string>(Files.Select(a => a.OldPath), StringComparer.OrdinalIgnoreCase);
        foreach (var action in Files)
        {
            /// If the destination exists AND it's not one of the files we are currently renaming
            if (File.Exists(action.NewPath) && !currentFiles.Contains(action.NewPath))
            {
                RenameAbort($"Collision detected: The file '{Path.GetFileName(action.NewPath)}' already exists in the destination.");
                return;
            }
        }

        var result = MessageBox.Show($"Are you sure you want to rename {Files.Count} files in the folder '{Folder}'?", TitleRename, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (result != MessageBoxResult.Yes)
        {
            RenameReset();
            return;
        }

        List<Exception> errors = new();
        bool criticalFailure = false;

        var info = new OperationInfo(Folder, []);
        await Task.Run(() =>
        {
            int.TryParse(DelayRename, out int delay);
            Thread.Sleep(delay);

            /// The undo stack remains our safety net if a move fails halfway through
            Stack<(string Current, string Original)> undoStack = new();
            try
            {
                int tempCounter = 1;

                /// 1. Rename temporarily to avoid collisions
                foreach (var action in Files)
                {
                    try
                    {
                        string directory = Path.GetDirectoryName(action.OldPath)!;

                        string tempName = $"{tempCounter++}.tmp";
                        string tempPath = Path.Combine(directory, tempName);

                        /// If by some freak accident 1.tmp already exists, we skip ahead
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
                    foreach (var action in Files)
                    {
                        try
                        {
                            File.Move(action.TempPath, action.NewPath);

                            /// Update the stack: the file is now at NewPath, needs to go back to OldPath on failure
                            /// We pop the old (temp->old) and push the new (new->old)
                            undoStack.Pop();
                            undoStack.Push((action.NewPath, action.OldPath));
                            info.Actions.Add((action.NewPath, action.OldPath));

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

        if (!criticalFailure)
        {
            RedoActions.Clear();
            UndoActions.Add(info);
        }

        Title = TitleRename;
        RenameReset();

        if (errors.Any())
        {
            LogMessages(errors);

            var message = $"One or more errors occurred while renaming.";
            message += LogEnable is true ? " Check log for more details!" : "";

            MessageBox.Show(message, TitleRename);
        }
        else
        {
            MessageBox.Show("Rename finished successfully!", TitleRename);
        }

        _ = Update();
    }

    private void RenameAbort(string message)
    {
        RenameReset();

        var error = new Exception(message);
        LogMessage(error);

        MessageBox.Show(error.Message, TitleRename, MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void RenameReset()
    {
        lock (Lock) { IsRenaming = false; IsBusy = false; }
    }


    private static string ToLetter(int n)
    {
        const int columns = 26;
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const int digitMaximum = 7; /// Enough for <see cref="Int32.MaxValue"/>

        /// Increment by 1 so that...

        ///  0 -> Processes as 1  -> Result "A"
        /// 25 -> Processes as 26 -> Result "Z"
        /// 26 -> Processes as 27 -> Result "AA"

        long current = (long)n + 1;

        /// Safety for negative
        if (current <= 0) current = 1;

        var result = new StringBuilder().Append(' ', digitMaximum);
        var offset = digitMaximum;

        while (current > 0)
        {
            /// Bijective conversion logic
            current--; /// Decrement to handle the "no zero digit" nature of letters
            result[--offset] = letters[(int)(current % columns)];
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


    private async Task Redo()
    {
        lock (Lock)
        {
            if (IsBusy || !RedoActions.Any()) return;

            IsBusy = true;
            IsRedoing = true;
        }

        var last = RedoActions.Last();
        last.Time = DateTime.Now;

        UndoActions.Add(last);
        RedoActions.Remove(last);

        await Task.Run(() =>
        {
            foreach (var operation in last.Actions)
            {
                try
                {
                    if (File.Exists(operation.OldPath))
                    {
                        if (!File.Exists(operation.NewPath))
                            File.Move(operation.OldPath, operation.NewPath);

                        else throw new Exception();
                    }
                    else throw new Exception();
                }
                catch (Exception) { } /// LogMessage(e);
            }
        });

        lock (Lock) { IsRedoing = false; IsBusy = false; }
        _ = Update();
    }

    private async Task Undo()
    {
        lock (Lock)
        {
            if (IsBusy || !UndoActions.Any()) return;

            IsBusy = true;
            IsUndoing = true;
        }

        var last = UndoActions.Last();
        last.Time = DateTime.Now;

        RedoActions.Add(last);
        UndoActions.Remove(last);

        await Task.Run(() =>
        {
            foreach (var operation in last.Actions)
            {
                try
                {
                    if (File.Exists(operation.NewPath))
                    {
                        if (!File.Exists(operation.OldPath))
                            File.Move(operation.NewPath, operation.OldPath);

                        else throw new Exception();
                    }
                    else throw new Exception();
                }
                catch (Exception) { } /// LogMessage(e);
            }
        });

        lock (Lock) { IsUndoing = false; IsBusy = false; }
        _ = Update();
    }


    private async Task Test()
    {
        var result = MessageBox.Show($"This will create a new folder with {TestFileCount} new files. Continue?", TitleRename, MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        lock (Lock)
        {
            if (IsBusy)
                return;

            IsBusy = true;
            IsTesting = true;
        }

        try
        {
            string fullPath = Path.Combine(TestTargetRoot, TestFolderName);
            int count = int.TryParse(TestFileCount, out int c) ? c : 10;
            int nameLen = int.TryParse(TestNameLength, out int l) ? l : 8;
            string ext = TestFileExtension.StartsWith(".") ? TestFileExtension : "." + TestFileExtension;

            await Task.Run(() =>
            {
                if (!Directory.Exists(fullPath))
                    Directory.CreateDirectory(fullPath);

                Random res = new Random();
                const string str = "abcdefghijklmnopqrstuvwxyz0123456789";

                for (int i = 0; i < count; i++)
                {
                    /// Generate random name
                    StringBuilder sb = new StringBuilder();
                    for (int j = 0; j < nameLen; j++)
                        sb.Append(str[res.Next(str.Length)]);

                    string fileName = sb.ToString() + ext;
                    string filePath = Path.Combine(fullPath, fileName);

                    /// Create empty dummy file
                    File.WriteAllText(filePath, "Test Data for Ion.Rename");
                }
            });

            if (TestAutoSelect)
                Folder = fullPath;

            MessageBox.Show($"Success! Created {count} files in {fullPath}", TitleRename);
        }
        catch (Exception x)
        {
            var message = $"Test failed: {x.Message}";

            LogMessage(new Exception(message));
            MessageBox.Show(message, TitleRename, MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            lock (Lock)
            {
                IsTesting = false;
                IsBusy = false;
            }
            _ = Update();
        }
    }

    internal async Task Update()
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
            if (!Validate())
                return;

            var folder = Folder;
            var includeSub = FolderSub == true;
            var actions = new List<Operation>();

            await Task.Run(() =>
            {
                int.TryParse(DelayUpdate, out int delay);
                Thread.Sleep(delay);

                try
                {
                    var all = GetFiles(folder, includeSub);
                    all = Filter(all);

                    var allGrouped = all.GroupBy(Path.GetDirectoryName);

                    var at = int.TryParse(FileNameIncrementAt, out var x) ? x : 1;

                    var iat = at;
                    var iby = int.TryParse(FileNameIncrementBy, out var y) ? y : 1;

                    var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (var group in allGrouped)
                    {
                        var folderFiles = group.OrderBy(f => f).ToList();
                        foreach (var file in folderFiles)
                        {
                            string oldExt = Path.GetExtension(file);
                            string newExt = GetFileExtension(file);

                            int current;
                            if (FileNameIncrementByExtension == true)
                            {
                                if (!counters.ContainsKey(oldExt)) counters[oldExt] = at;
                                current = counters[oldExt];
                                counters[oldExt] += iby;
                            }
                            else
                            {
                                current = iat;
                                iat += iby;
                            }

                            string newName = GetFileName(current) + newExt;
                            string newPath = Path.Combine(Path.GetDirectoryName(file)!, newName);

                            actions.Add(new Operation
                            {
                                OldPath = file, OldName = Path.GetFileName(file),
                                NewName = newName, NewPath = newPath,
                                RelativeFolder = Path.GetRelativePath(folder, Path.GetDirectoryName(file)!)
                            });
                        }
                    }
                }
                catch (Exception e)
                {
                    LogMessage(new Exception(string.Format(Messages.UpdateFailed, e.Message)));
                }
            });

            Files = new ObservableCollection<Operation>(actions);

            FileCount = Files.Count;
            FileCountText = FileCount == 1 ? "1 file" : $"{FileCount} files";
        }
        finally
        {
            bool runAgain = false;
            lock (Lock)
            {
                IsUpdating = false;
                IsBusy = false;

                if (Dirty)
                {
                    runAgain = true;
                    Dirty = false;
                }
            }

            if (runAgain)
                Dispatcher.BeginInvoke(new Action(() => Update()));
        }
    }

    bool Validate()
    {
        if (!int.TryParse(DelayRename, out int _))
        {
            LogMessage(new Exception(Messages.DelayRenameNotValid));
            return false;
        }

        if (!int.TryParse(DelayUpdate, out int _))
        {
            LogMessage(new Exception(Messages.DelayUpdateNotValid));
            return false;
        }

        if (FileExtensionOverride == true)
        {
            if (!string.IsNullOrEmpty(FileExtensionOverrideNew))
            {
                LogMessage(new Exception(Messages.FileExtensionOverride));
                if (FileExtensionOverrideNew.Any(x => Path.GetInvalidFileNameChars().Contains(x)))
                {
                    LogMessage(new Exception(Messages.FileExtensionOverrideNotValid));
                    return false;
                }
            }
        }

        if (string.IsNullOrEmpty(FileNameFormat))
        {
            LogMessage(new Exception(Messages.FileNameFormatNotSpecified));
            return false;
        }

        if (!FileNameFormat.Contains(FileNameFormatSymbol))
        {
            LogMessage(new Exception(Messages.FileNameFormatSymbolNotSpecified));
            return false;
        }

        if (FileNameFormat.Replace(FileNameFormatSymbol, "").Any(c => Path.GetInvalidFileNameChars().Contains(c)))
        {
            LogMessage(new Exception(Messages.FileNameFormatNotValid));
            return false;
        }

        if (!int.TryParse(FileNameIncrementAt, out int _))
        {
            LogMessage(new Exception(Messages.FileNameIncrementAtNotValid));
            return false;
        }

        if (!int.TryParse(FileNameIncrementBy, out int _))
        {
            LogMessage(new Exception(Messages.FileNameIncrementByNotValid));
            return false;
        }

        if (FileNameNumberPaddingCharacter is null || FileNameNumberPaddingCharacter.Length <= 0)
        {
            LogMessage(new Exception(Messages.FileNameNumberPaddingCharacterNotSpecified));
            return false;
        }
        else if (FileNameNumberPaddingCharacter.Any(x => Path.GetInvalidFileNameChars().Contains(x)) == true)
        {
            LogMessage(new Exception(Messages.FileNameNumberPaddingCharacterNotValid));
            return false;
        }

        if (string.IsNullOrWhiteSpace(Folder))
            return false;

        if (!Directory.Exists(Folder))
        {
            LogMessage(new Exception(string.Format(Messages.FolderExists, Folder)));
            return false;
        }

        if (false && !IsFolderChildOf(Folder, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)))
        {
            LogMessage(new Exception(Messages.FolderNotValid));
            return false;
        }

        if (!IsFolderSafe(Folder))
        {
            LogMessage(new Exception($"Folder base path too long. Windows may fail to rename files in folder '{Folder}'."));
            return false;
        }

        return true;
    }

    #endregion

    /// <see cref="EventArgs"/>
    #region

    private void OnLogChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        LogCount = Log.Count;
    }

    private void OnUpdated(object sender, EventArgs e) => _ = Update();

    #endregion

    /// <see cref="DragEventArgs"/>
    #region

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            IsDraggingOver = true;
            IsDraggingOverFolder = e.Data.GetData(DataFormats.FileDrop) is string[] x && x.Length > 0 ? x[0] : "";

            e.Effects = DragDropEffects.Copy;
        }
    }

    private void DragDropOverlay_DragLeave(object sender, DragEventArgs e)
    {
        IsDraggingOver = false;
    }

    private void DragDropOverlay_Drop(object sender, DragEventArgs e)
    {
        IsDraggingOver = false;

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (paths.Length > 0)
            {
                string selectedPath = paths[0];

                /// Logic: If it's a file, get the directory. If it's a folder, use it.
                string targetFolder = Directory.Exists(selectedPath)
                ? selectedPath
                : Path.GetDirectoryName(selectedPath)!;

                /// PREEMPTIVE PATH CHECK (Issue #4)
                if (targetFolder.Length > 240)
                {
                    MessageBox.Show("Warning: This path is approaching the Windows limit. Some renames may fail.",
                        "Path Length Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                Folder = targetFolder;
            }
        }
    }

    #endregion

    /// <see cref="RoutedEventArgs"/>
    #region

    private void OnBrowse(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "Select folder..." };
        if (dialog.ShowDialog() == true)
            Folder = dialog.FolderName;
    }

    private void OnBrowseTestRoot(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog { Title = "Select Root for Test Folder" };
        if (dialog.ShowDialog() == true)
            TestTargetRoot = dialog.FolderName;
    }

    private void OnClearHistory(object sender, RoutedEventArgs e)
    {
        lock (Lock)
        {
            if (IsBusy) return;
            IsBusy = true;
        }
        UndoActions.Clear();
        RedoActions.Clear();
        lock (Lock) { IsBusy = false; }
    }

    private void OnFilterAdd(object sender, RoutedEventArgs e) => Filters.Add(new FilterItem(this));

    private void OnFiltersCleared(object sender, RoutedEventArgs e) => Filters.Clear();

    private void OnFilterRemoved(object sender, RoutedEventArgs e)
    {
        if (sender is Button { DataContext: FilterItem item }) 
            Filters.Remove(item);
    }

    private void OnFilterRemovedSelected(object sender, RoutedEventArgs e)
    {
        var selected = FilterListBox.SelectedItems.Cast<FilterItem>().ToList();
        foreach (var item in selected) Filters.Remove(item);
    }

    private void OnRedo(object sender, RoutedEventArgs e) => _ = Redo();

    private void OnTest(object sender, RoutedEventArgs e) => _ = Test();

    private void OnUndo(object sender, RoutedEventArgs e) => _ = Undo();

    private void OnUpdate(object sender, RoutedEventArgs e) => _ = Update();

    private void OnRenamed(object sender, RoutedEventArgs e) => _ = Rename();

    #endregion

    /// <see cref="INotifyPropertyChanged"/>
    #region

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    #endregion
}

#endregion

/// <see cref="FilterItem"/>
#region

public class FilterItem : INotifyPropertyChanged
{
    private readonly MainWindow _parent;
    public FilterItem(MainWindow parent) => _parent = parent;

    public string[] AllAttributes => [
        "File name", "Extension", "Folder Path",
        "Size (Bytes)",
        "Date Created", "Date Modified", "Date Accessed",
        "Is ReadOnly", "Is Hidden", "Is System", "Is Archive"
    ];

    public string[] StringModes => ["Contains", "Equals", "Starts with", "Ends with", "Regex"];
    public string[] NumericModes => ["Greater than", "Less than", "Exactly"];
    public string[] DateModes => ["After", "Before", "On"];

    private string _targetAttribute = "File name";
    public string TargetAttribute
    {
        get => _targetAttribute;
        set
        {
            _targetAttribute = value;
            UpdateType();
            OnPropertyChanged();
        }
    }

    private FilterType _currentType = FilterType.String;
    public FilterType CurrentType { get => _currentType; private set { _currentType = value; OnPropertyChanged(); } }

    private string _selectedMode = "Contains";
    public string SelectedMode { get => _selectedMode; set { _selectedMode = value; OnPropertyChanged(); } }

    private string _stringValue = "";
    public string StringValue { get => _stringValue; set { _stringValue = value; OnPropertyChanged(); } }

    private long _longValue = 0;
    public long LongValue { get => _longValue; set { _longValue = value; OnPropertyChanged(); } }

    private DateTime _dateValue = DateTime.Now;
    public DateTime DateValue { get => _dateValue; set { _dateValue = value; OnPropertyChanged(); } }

    private bool _boolValue = true;
    public bool BoolValue { get => _boolValue; set { _boolValue = value; OnPropertyChanged(); } }

    private void UpdateType()
    {
        CurrentType = TargetAttribute switch
        {
            "File name" or "Extension" or "Folder Path" => FilterType.String,
            "Size (Bytes)" => FilterType.Numeric,
            "Date Created" or "Date Modified" or "Date Accessed" => FilterType.DateTime,
            _ => FilterType.Boolean
        };

        /// Reset default mode based on type
        SelectedMode = CurrentType switch
        {
            FilterType.String => "Contains",
            FilterType.Numeric => "Greater Than",
            FilterType.DateTime => "After",
            _ => ""
        };
    }

    public bool IsMatch(FileInfo file)
    {
        return CurrentType switch
        {
            FilterType.String => ApplyString(file),
            FilterType.Numeric => ApplyNumeric(file),
            FilterType.DateTime => ApplyDate(file),
            FilterType.Boolean => ApplyBool(file),
            _ => true
        };
    }

    private bool ApplyString(FileInfo file)
    {
        string val = TargetAttribute switch
        {
            "File name" => Path.GetFileNameWithoutExtension(file.Name),
            "Extension" => file.Extension,
            _ => file.DirectoryName ?? ""
        };
        if (SelectedMode == "Regex")
        {
            try
            {
                return System.Text.RegularExpressions.Regex.IsMatch(val, StringValue, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            }
            catch { return false; } /// Invalid Regex
        }
        return SelectedMode switch
        {
            "Contains" => val.Contains(StringValue, StringComparison.OrdinalIgnoreCase),
            "Starts with" => val.StartsWith(StringValue, StringComparison.OrdinalIgnoreCase),
            "Ends with" => val.EndsWith(StringValue, StringComparison.OrdinalIgnoreCase),
            _ => val.Equals(StringValue, StringComparison.OrdinalIgnoreCase)
        };
    }

    private bool ApplyNumeric(FileInfo file) => SelectedMode switch
    {
        "Greater Than" => file.Length > LongValue,
        "Less Than" => file.Length < LongValue,
        _ => file.Length == LongValue
    };

    private bool ApplyDate(FileInfo file)
    {
        DateTime dt = TargetAttribute == "Date Created" ? file.CreationTime : file.LastWriteTime;
        return SelectedMode switch
        {
            "After" => dt > DateValue,
            "Before" => dt < DateValue,
            _ => dt.Date == DateValue.Date
        };
    }

    private bool ApplyBool(FileInfo file)
    {
        bool isTarget = TargetAttribute switch
        {
            "Is ReadOnly" => file.IsReadOnly,
            "Is Hidden" => (file.Attributes & FileAttributes.Hidden) != 0,
            _ => (file.Attributes & FileAttributes.System) != 0
        };
        return isTarget == BoolValue;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        _parent.Update();
    }
}

#endregion

/// <see cref="FilterType"/>
#region

public enum FilterType { String, Numeric, DateTime, Boolean }

public class FilterTypeToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is FilterType current && parameter is string target)
            return current.ToString() == target ? Visibility.Visible : Visibility.Collapsed;

        return Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
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

public class NotNullConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is not null;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

public class NotZeroConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is Byte a)
            return a != 0;

        if (value is Decimal b)
            return b != 0;

        if (value is Double c)
            return c != 0;

        if (value is Single d)
            return d != 0;


        if (value is Int16 x)
            return x != 0;

        if (value is Int32 y)
            return y != 0;

        if (value is Int64 z)
            return z != 0;

        return Binding.DoNothing;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}

public class ThumbnailConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is string i ? Thumbnail.GetIcon(i, ThumbnailSize.Jumbo) : Binding.DoNothing;

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

/// <see cref="Message"/>
#region

public class Message(Exception Error)
{
    public Exception Error { get; set; } = Error;

    public DateTime Time { get; private set; } = DateTime.Now;
}

public class Messages
{
    public const string DelayRenameNotValid = "Delay rename not valid.";

    public const string DelayUpdateNotValid = "Delay update not valid.";

    public const string FolderExists = "Folder '{0}' does not exist.";

    public const string FolderNotSpecified = $"Folder not specified.";

    public const string FolderNotValid = "Folder is not a subfolder of the current user folder.";

    public const string FileExtensionOverride = "Changing a file extension may cause a file to become unreadable in some contexts.";
    
    public const string FileExtensionOverrideNotValid = "File extension override not valid.";

    public const string FileNameFormatNotSpecified = "File name format not specified.";

    public const string FileNameFormatSymbolNotSpecified = $"File name format must include '{MainWindow.FileNameFormatSymbol}'.";

    public const string FileNameFormatNotValid = "File name format not valid.";

    public const string FileNameIncrementAtNotValid = "File name increment at not valid.";

    public const string FileNameIncrementByNotValid = "File name increment by not valid.";

    public const string FileNameNumberPaddingCharacterNotSpecified = "File name number padding character not specified.";

    public const string FileNameNumberPaddingCharacterNotValid = "File name number padding character not valid.";

    public const string NumberNegative = "Must not be negative.";

    public const string NumberNotValid = "Not a valid number.";
    
    public const string Required = "This is required.";

    public const string UpdateFailed = "Update failed: {0}";
}

#endregion

/// <see cref="Operation"/>
#region

public class Operation : INotifyPropertyChanged
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

/// <see cref="OperationInfo"/>
#region

public class OperationInfo(string Folder, IEnumerable<(string NewPath, string OldPath)> Actions) : INotifyPropertyChanged
{
    public List<(string NewPath, string OldPath)> Actions { get => field; private set => field = value; } = new(Actions);

    public string Folder
    {
        get => field;
        set
        {
            field = value;
            OnPropertyChanged();
        }
    } = "";

    public DateTime Time { get => field; set { field = value; OnPropertyChanged(); } } = DateTime.Now;

    /// <see cref="INotifyPropertyChanged"/>
    #region

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    #endregion
}

#endregion

/// <see cref="Preset"/>
#region

public static class Preset
{
    public static readonly DependencyProperty IdProperty = DependencyProperty.RegisterAttached("Id", typeof(string), typeof(Preset), new PropertyMetadata(null));
    public static void SetId(DependencyObject obj, string value) => obj.SetValue(IdProperty, value);
    public static string GetId(DependencyObject obj) => (string)obj.GetValue(IdProperty);
    
    public static RelayCommand CopyCommand { get; } = new RelayCommand(i =>
    {
        if (i is TextBox j)
            Clipboard.SetText(j.Text);
    });

    public static RelayCommand OpenCommand { get; } = new RelayCommand(i =>
    {
        if (i is TextBox j)
        {
            var id = GetId(j);
            var window = new PresetWindow(id);
            if (window.ShowDialog() == true && window.SelectedValue != null)
            {
                j.Text = window.SelectedValue.Value;
                j.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            }
        }
    });
}

#endregion

/// <see cref="PresetItem"/>
#region

public class PresetItem(string value) : INotifyPropertyChanged
{
    private string _val = value;
    public string Value
    {
        get => _val;
        set { _val = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

#endregion

/// <see cref="PresetManager"/>
#region

public static class PresetManager
{
    private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "presets.json");
    
    /// Key: Preset ID (e.g., "FileNameFormat"), Value: List of strings
    public static Dictionary<string, ObservableCollection<PresetItem>> Store { get; private set; } = new();

    public static void Load()
    {
        if (File.Exists(FilePath))
        {
            try {
                var json = File.ReadAllText(FilePath);
                Store = JsonSerializer.Deserialize<Dictionary<string, ObservableCollection<PresetItem>>>(json) ?? new();
            } catch { Store = new(); }
        }
    }

    public static void Save()
    {
        var json = JsonSerializer.Serialize(Store, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(FilePath, json);
    }

    public static ObservableCollection<PresetItem> GetOrCreate(string id)
    {
        if (!Store.ContainsKey(id)) Store[id] = new ObservableCollection<PresetItem>();
        return Store[id];
    }
}

#endregion

/// <see cref="RelayCommand"/>
#region

public class RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null) : ICommand
{
    private readonly Action<object?> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
    private readonly Predicate<object?>? _canExecute = canExecute;

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) => _canExecute == null || _canExecute(parameter);

    public void Execute(object? parameter) => _execute(parameter);
}

#endregion

/// <see cref="ValidationRule"/>
#region

public class RuleFileExtension : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        if (string.IsNullOrEmpty(text))
            return ValidationResult.ValidResult;

        return new ValidationResult(false, Messages.FileExtensionOverride);
    }
}

public class RuleFileName : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;
        if (text?.Any(c => Path.GetInvalidFileNameChars().Contains(c)) == true)
            return new ValidationResult(false, "Invalid characters.");

        return ValidationResult.ValidResult;
    }
}

public class RuleFileNameFormat : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        if (string.IsNullOrEmpty(text))
            return new ValidationResult(false, Messages.FileNameFormatNotSpecified);

        if (!text.Contains(MainWindow.FileNameFormatSymbol))
            return new ValidationResult(false, Messages.FileNameFormatSymbolNotSpecified);

        text = text.Replace(MainWindow.FileNameFormatSymbol, "");
        if (text.Any(c => Path.GetInvalidFileNameChars().Contains(c)))
            return new ValidationResult(false, Messages.FileNameFormatNotValid);

        return ValidationResult.ValidResult;
    }
}

public class RuleFolder : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;
        if (string.IsNullOrEmpty(text) || Directory.Exists(text))
            return ValidationResult.ValidResult;

        return new ValidationResult(false, "Folder doesn't exist!");
    }
}

public class RuleInt32 : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        var result = int.TryParse(text, out int _);
        if (result)
            return ValidationResult.ValidResult;

        return new ValidationResult(false, Messages.NumberNotValid);
    }
}

public class RuleInt32Negative : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        var result = int.TryParse(text, out int x);
        if (result && x >= 0)
            return ValidationResult.ValidResult;

        return new ValidationResult(false, Messages.NumberNegative);
    }
}

public class RuleInt64 : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;

        var result = long.TryParse(text, out long _);
        if (result)
            return ValidationResult.ValidResult;

        return new ValidationResult(false, Messages.NumberNotValid);
    }
}

public class RuleMinimum : ValidationRule
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

public class RuleRequire : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        string? text = value as string;
        if (text is null || text.Length == 0)
            return new ValidationResult(false, Messages.Required);

        return ValidationResult.ValidResult;
    }
}

#endregion

/// <see cref="XComboBox"/>
#region

public static class XComboBox
{
    public static readonly DependencyProperty HideArrowProperty = DependencyProperty.RegisterAttached("HideArrow", typeof(bool), typeof(XComboBox), new PropertyMetadata(false, OnHideArrowChanged));
    public static bool GetHideArrow(DependencyObject obj) => (bool)obj.GetValue(HideArrowProperty);
    public static void SetHideArrow(DependencyObject obj, bool value) => obj.SetValue(HideArrowProperty, value);
    private static void OnHideArrowChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is ComboBox cb && (bool)e.NewValue)
        {
            cb.Loaded += (s, arg) =>
            {
                cb.ApplyTemplate();

                var toggleButton = FindVisualChild<ToggleButton>(cb);
                if (toggleButton != null)
                {
                    toggleButton.ApplyTemplate();

                    var arrow = FindVisualChild<System.Windows.Shapes.Path>(toggleButton);
                    if (arrow != null) arrow.Visibility = Visibility.Collapsed;
                }
            };
        }
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T t) return t;
            var result = FindVisualChild<T>(child);
            if (result != null) return result;
        }
        return null;
    }
}

#endregion
