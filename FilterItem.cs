using System.ComponentModel;

namespace OutlookEventForwarder;

public class FilterItem : INotifyPropertyChanged
{
    private bool _isChecked = true;

    public string Value { get; set; } = "";

    public bool IsChecked
    {
        get => _isChecked;
        set
        {
            _isChecked = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}
