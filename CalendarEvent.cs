using System.ComponentModel;

namespace OutlookEventForwarder;

public class CalendarEvent : INotifyPropertyChanged
{
    private bool _isSelected;

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            _isSelected = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
        }
    }

    public string Subject { get; set; } = "";
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string Location { get; set; } = "";
    public string Organizer { get; set; } = "";
    public string EntryId { get; set; } = "";

    public event PropertyChangedEventHandler? PropertyChanged;
}
