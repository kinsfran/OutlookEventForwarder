using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

namespace OutlookEventForwarder;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<CalendarEvent> _events = new();
    private readonly List<CalendarEvent> _allEvents = new();
    private readonly ObservableCollection<string> _recipients = new();
    private readonly Dictionary<string, HashSet<string>> _activeFilters = new();
    private readonly ObservableCollection<FilterItem> _currentFilterItems = new();
    private string? _currentFilterColumn;
    private int _lastClickedIndex = -1;
    private OutlookApp? _outlookApp;

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OutlookEventForwarder", "settings.json");

    public MainWindow()
    {
        InitializeComponent();

        DateFrom.SelectedDate = DateTime.Today;
        DateTo.SelectedDate = DateTime.Today.AddDays(30);

        EventsGrid.ItemsSource = _events;
        RecipientList.ItemsSource = _recipients;
        FilterItems.ItemsSource = _currentFilterItems;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        LoadColumnWidths();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        SaveColumnWidths();
    }

    private OutlookApp GetOutlookApp()
    {
        if (_outlookApp != null) return _outlookApp;
        _outlookApp = new OutlookApp();
        return _outlookApp;
    }

    private void BtnLoadEvents_Click(object sender, RoutedEventArgs e)
    {
        if (DateFrom.SelectedDate == null || DateTo.SelectedDate == null)
        {
            MessageBox.Show("Please select both From and To dates.", "Missing Dates",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Items? calendarItems = null;
        MAPIFolder? calendarFolder = null;
        NameSpace? ns = null;

        try
        {
            _allEvents.Clear();
            _events.Clear();
            _activeFilters.Clear();
            var app = GetOutlookApp();
            ns = app.GetNamespace("MAPI");
            calendarFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            calendarItems = calendarFolder.Items;

            calendarItems.Sort("[Start]", false);
            calendarItems.IncludeRecurrences = false;

            var fromDate = DateFrom.SelectedDate.Value;
            var toDate = DateTo.SelectedDate.Value.AddDays(1);
            var from = fromDate.ToString("g");
            var to = toDate.ToString("g");
            var filter = $"[Start] >= '{from}' AND [End] <= '{to}'";
            var subjectFilter = SubjectFilter.Text.Trim();

            object? item = calendarItems.Find(filter);
            while (item != null)
            {
                try
                {
                    if (item is AppointmentItem appt
                        && appt.Start >= fromDate && appt.Start < toDate)
                    {
                        var subject = appt.Subject;
                        if (string.IsNullOrEmpty(subjectFilter) ||
                            (subject != null && subject.Contains(subjectFilter, StringComparison.OrdinalIgnoreCase)))
                        {
                            var ev = new CalendarEvent
                            {
                                Subject = subject ?? "(No Subject)",
                                Start = appt.Start,
                                End = appt.End,
                                Location = appt.Location ?? "",
                                Organizer = appt.Organizer ?? "",
                                EntryId = appt.EntryID
                            };
                            _allEvents.Add(ev);
                            _events.Add(ev);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(item);
                }

                item = calendarItems.FindNext();
            }

            PlaceholderText.Visibility = _events.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
            if (_events.Count == 0)
                PlaceholderText.Text = "No events found for the selected date range.";

            UpdateStatusText();
        }
        catch (System.Exception ex)
        {
            App.LogError("BtnLoadEvents_Click", ex);
            MessageBox.Show(
                $"Failed to load events (logged to Desktop):\n\n{ex.GetType().Name}: {ex.Message}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (calendarItems != null) Marshal.ReleaseComObject(calendarItems);
            if (calendarFolder != null) Marshal.ReleaseComObject(calendarFolder);
            if (ns != null) Marshal.ReleaseComObject(ns);
        }
    }

    private string GetColumnValue(CalendarEvent ev, string column) => column switch
    {
        "Subject" => ev.Subject,
        "Start" => ev.StartDisplay,
        "End" => ev.EndDisplay,
        "Location" => ev.Location,
        "Organizer" => ev.Organizer,
        _ => ""
    };

    private void FilterButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not string columnName) return;

        _currentFilterColumn = columnName;
        _currentFilterItems.Clear();

        var values = _allEvents
            .Select(ev => GetColumnValue(ev, columnName))
            .Distinct()
            .OrderBy(v => v)
            .ToList();

        _activeFilters.TryGetValue(columnName, out var checkedValues);

        foreach (var val in values)
        {
            _currentFilterItems.Add(new FilterItem
            {
                Value = string.IsNullOrEmpty(val) ? "(empty)" : val,
                IsChecked = checkedValues == null || checkedValues.Contains(string.IsNullOrEmpty(val) ? "(empty)" : val)
            });
        }

        FilterPopup.PlacementTarget = btn;
        FilterPopup.IsOpen = true;
    }

    private void FilterItem_Changed(object sender, RoutedEventArgs e)
    {
        ApplyCurrentFilter();
    }

    private void FilterSelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _currentFilterItems)
            item.IsChecked = true;
        ApplyCurrentFilter();
    }

    private void FilterClearAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _currentFilterItems)
            item.IsChecked = false;
        ApplyCurrentFilter();
    }

    private void ApplyCurrentFilter()
    {
        if (_currentFilterColumn == null) return;

        var checkedValues = _currentFilterItems
            .Where(f => f.IsChecked)
            .Select(f => f.Value)
            .ToHashSet();

        if (checkedValues.Count == _currentFilterItems.Count)
            _activeFilters.Remove(_currentFilterColumn);
        else
            _activeFilters[_currentFilterColumn] = checkedValues;

        ApplyAllFilters();
    }

    private void ApplyAllFilters()
    {
        _events.Clear();

        foreach (var ev in _allEvents)
        {
            bool passes = true;
            foreach (var (column, allowed) in _activeFilters)
            {
                var val = GetColumnValue(ev, column);
                var displayVal = string.IsNullOrEmpty(val) ? "(empty)" : val;
                if (!allowed.Contains(displayVal))
                {
                    passes = false;
                    break;
                }
            }
            if (passes)
                _events.Add(ev);
        }

        PlaceholderText.Visibility = _events.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
        BtnClearFilters.Visibility = _activeFilters.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        UpdateStatusText();
    }

    private void BtnClearFilters_Click(object sender, RoutedEventArgs e)
    {
        _activeFilters.Clear();
        ApplyAllFilters();
    }

    private void UpdateStatusText()
    {
        var selectedCount = _events.Count(ev => ev.IsSelected);
        var filterInfo = _activeFilters.Count > 0 ? $" ({_activeFilters.Count} filter(s) active)" : "";
        var selectedInfo = selectedCount > 0 ? $" | {selectedCount} selected" : "";
        StatusText.Text = $"Showing {_events.Count} of {_allEvents.Count} event(s){filterInfo}{selectedInfo}";
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox cb)
        {
            var selected = cb.IsChecked == true;
            foreach (var ev in _events)
                ev.IsSelected = selected;
        }
    }

    private void AddRecipient()
    {
        var input = RecipientInput.Text.Trim();
        if (string.IsNullOrEmpty(input)) return;

        var emails = input.Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var invalid = new List<string>();

        foreach (var email in emails)
        {
            if (!email.Contains('@'))
            {
                invalid.Add(email);
                continue;
            }
            if (!_recipients.Contains(email))
                _recipients.Add(email);
        }

        if (invalid.Count > 0)
            MessageBox.Show($"Skipped invalid email(s):\n{string.Join("\n", invalid)}", "Invalid Email",
                MessageBoxButton.OK, MessageBoxImage.Warning);

        RecipientInput.Text = "";
        RecipientInput.Focus();
    }

    private void RecipientInput_TextChanged(object sender, TextChangedEventArgs e)
    {
        RecipientPlaceholder.Visibility = string.IsNullOrEmpty(RecipientInput.Text)
            ? Visibility.Visible : Visibility.Collapsed;
    }

    private void BtnAddRecipient_Click(object sender, RoutedEventArgs e) => AddRecipient();

    private void RecipientInput_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddRecipient();
            e.Handled = true;
        }
    }

    private void RemoveRecipient_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string email)
            _recipients.Remove(email);
    }

    private void BtnForward_Click(object sender, RoutedEventArgs e)
    {
        var selectedEvents = _events.Where(ev => ev.IsSelected).ToList();
        if (selectedEvents.Count == 0)
        {
            MessageBox.Show("Please select at least one event to forward.", "No Events Selected",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (_recipients.Count == 0)
        {
            MessageBox.Show("Please add at least one recipient.", "No Recipients",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            $"Forward {selectedEvents.Count} event(s) to {_recipients.Count} recipient(s)?\n\n" +
            $"This will send {selectedEvents.Count} forwarding email(s), each to all {_recipients.Count} recipient(s).",
            "Confirm Forward", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes) return;

        NameSpace? ns = null;
        try
        {
            BtnForward.IsEnabled = false;
            var app = GetOutlookApp();
            ns = app.GetNamespace("MAPI");
            int sent = 0;
            int failed = 0;

            foreach (var ev in selectedEvents)
            {
                AppointmentItem? appt = null;
                MailItem? mail = null;
                try
                {
                    appt = (AppointmentItem)ns.GetItemFromID(ev.EntryId);
                    mail = (MailItem)app.CreateItem(OlItemType.olMailItem);
                    mail.Subject = "FW: " + appt.Subject;

                    var body = new StringBuilder();
                    body.AppendLine("---------- Forwarded event ----------");
                    body.AppendLine($"Subject: {appt.Subject}");
                    body.AppendLine($"When: {appt.Start:dddd, MMMM d, yyyy h:mm tt} - {appt.End:h:mm tt}");
                    if (!string.IsNullOrEmpty(appt.Location))
                        body.AppendLine($"Location: {appt.Location}");
                    body.AppendLine($"Organizer: {appt.Organizer}");
                    if (!string.IsNullOrEmpty(appt.Body))
                    {
                        body.AppendLine();
                        body.AppendLine(appt.Body);
                    }
                    mail.Body = body.ToString();

                    foreach (var recipient in _recipients)
                        mail.Recipients.Add(recipient);
                    mail.Recipients.ResolveAll();
                    mail.Send();
                    sent++;
                }
                catch (System.Exception ex)
                {
                    App.LogError($"Forward event '{ev.Subject}'", ex);
                    failed++;
                }
                finally
                {
                    if (mail != null) Marshal.ReleaseComObject(mail);
                    if (appt != null) Marshal.ReleaseComObject(appt);
                }

                StatusText.Text = $"Forwarding... {sent + failed}/{selectedEvents.Count}";
            }

            var message = $"Done! Forwarded {sent} event(s) to {_recipients.Count} recipient(s).";
            if (failed > 0)
                message += $"\n{failed} event(s) failed to forward.";

            StatusText.Text = message;
            MessageBox.Show(message, "Complete", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (System.Exception ex)
        {
            App.LogError("BtnForward_Click", ex);
            MessageBox.Show($"Error during forwarding:\n\n{ex.Message}",
                "Forward Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (ns != null) Marshal.ReleaseComObject(ns);
            BtnForward.IsEnabled = true;
        }
    }

    private void EventsGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var dep = (DependencyObject)e.OriginalSource;
        while (dep != null && dep is not DataGridRow)
            dep = System.Windows.Media.VisualTreeHelper.GetParent(dep);

        if (dep is not DataGridRow row || row.Item is not CalendarEvent clickedEvent) return;

        var clickedIndex = _events.IndexOf(clickedEvent);
        if (clickedIndex < 0) return;

        // Check if clicking on the checkbox column — let the checkbox handle it
        var cell = FindParent<DataGridCell>(e.OriginalSource as DependencyObject);
        if (cell != null && cell.Column == EventsGrid.Columns[0])
        {
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && _lastClickedIndex >= 0)
            {
                var start = Math.Min(_lastClickedIndex, clickedIndex);
                var end = Math.Max(_lastClickedIndex, clickedIndex);
                var newState = !clickedEvent.IsSelected;
                for (int i = start; i <= end; i++)
                    _events[i].IsSelected = newState;

                UpdateStatusText();
                e.Handled = true;
                return;
            }

            _lastClickedIndex = clickedIndex;
            Dispatcher.BeginInvoke(() => UpdateStatusText(), System.Windows.Threading.DispatcherPriority.Background);
            return;
        }
    }

    private static T? FindParent<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child != null)
        {
            if (child is T result) return result;
            child = System.Windows.Media.VisualTreeHelper.GetParent(child);
        }
        return null;
    }

    private void EventsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (EventsGrid.CurrentItem is not CalendarEvent ev) return;

        NameSpace? ns = null;
        AppointmentItem? appt = null;
        try
        {
            var app = GetOutlookApp();
            ns = app.GetNamespace("MAPI");
            appt = (AppointmentItem)ns.GetItemFromID(ev.EntryId);
            appt.Display(false);
        }
        catch (System.Exception ex)
        {
            App.LogError("DoubleClick_Open", ex);
        }
        finally
        {
            if (appt != null) Marshal.ReleaseComObject(appt);
            if (ns != null) Marshal.ReleaseComObject(ns);
        }
    }

    private void SaveColumnWidths()
    {
        try
        {
            var widths = new Dictionary<string, double>();
            foreach (var col in EventsGrid.Columns)
            {
                var header = col.Header?.ToString();
                if (header != null)
                    widths[header] = col.ActualWidth;
            }

            var dir = Path.GetDirectoryName(SettingsPath)!;
            Directory.CreateDirectory(dir);
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(widths));
        }
        catch { }
    }

    private void LoadColumnWidths()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return;

            var json = File.ReadAllText(SettingsPath);
            var widths = JsonSerializer.Deserialize<Dictionary<string, double>>(json);
            if (widths == null) return;

            foreach (var col in EventsGrid.Columns)
            {
                var header = col.Header?.ToString();
                if (header != null && widths.TryGetValue(header, out var width))
                    col.Width = new DataGridLength(width);
            }
        }
        catch { }
    }
}
