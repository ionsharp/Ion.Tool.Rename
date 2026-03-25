using System.Collections.ObjectModel;
using System.Windows;

namespace Ion.Tool.Rename;

public partial class PresetWindow : Window
{
    public ObservableCollection<PresetItem> Items { get; set; }

    public PresetItem SelectedValue { get; private set; }

    public PresetWindow(string id)
    {
        InitializeComponent();
        Items = PresetManager.GetOrCreate(id);
        DataContext = this;
    }

    private void OnAdd(object sender, RoutedEventArgs e)
    {
        // Simple InputBox or prompt
        Items.Add(new PresetItem(""));
    }

    private void OnClear(object sender, RoutedEventArgs e)
    {
        Items.Clear();
    }
    
    private void OnRemove(object sender, RoutedEventArgs e)
    {
        if (PresetList.SelectedItem is PresetItem s) Items.Remove(s);
    }

    private void OnSelect(object sender, RoutedEventArgs e)
    {
        SelectedValue = PresetList.SelectedItem as PresetItem;
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        SelectedValue = null;
        DialogResult = false;

        Close();
    }
}