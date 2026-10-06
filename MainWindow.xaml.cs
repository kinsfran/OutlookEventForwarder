using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.Outlook;
using OutlookApplication = Microsoft.Office.Interop.Outlook.Application;

namespace OutlookEventForwarder;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<CalendarEvent> _events = new();
    private readonly ObservableCollection<string> _recipients = new();
    private OutlookApplication? _outlookApp;

    public MainWindow()
    {
        InitializeComponent();

        DateFrom.SelectedDate = DateTime.Today;
        DateTo.SelectedDate = DateTime.Today.AddDays(30);

        EventsGrid.ItemsSource = _events;
        RecipientList.ItemsSource = _recipients;
    }

    private OutlookApplication GetOutlookApp()
    {
        _outlookApp ??= new OutlookApplication();
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

        try
        {
            _events.Clear();
            var app = GetOutlookApp();
            var ns = app.GetNamespace("MAPI");
            var calendarFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            var items = calendarFolder.Items;

            items.Sort("[Start]", false);
            items.IncludeRecurrences = true;

            var from = DateFrom.SelectedDate.Value.ToString("g");
            var to = DateTo.SelectedDate.Value.AddDays(1).ToString("g");
            var filter = $"[Start] >= '{from}' AND [Start] < '{to}'";

            var restricted = items.Restrict(filter);
            var subjectFilter = SubjectFilter.Text.Trim();

            foreach (object item in restricted)
            {
                if (item is not AppointmentItem appt) continue;

                if (!string.IsNullOrEmpty(subjectFilter) &&
                    (appt.Subject == null || !appt.Subject.Contains(subjectFilter, StringComparison.OrdinalIgnoreCase)))
                    continue;

                _events.Add(new CalendarEvent
                {
                    Subject = appt.Subject ?? "(No Subject)",
                    Start = appt.Start,
                    End = appt.End,
                    Location = appt.Location ?? "",
                    Organizer = appt.Organizer ?? "",
                    EntryId = appt.EntryID
                });
            }

            PlaceholderText.Visibility = _events.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
            if (_events.Count == 0)
                PlaceholderText.Text = "No events found for the selected date range.";

            StatusText.Text = $"Loaded {_events.Count} event(s).";
        }
        catch (System.Exception ex)
        {
            MessageBox.Show($"Failed to load events from Outlook:\n\n{ex.Message}",
                "Outlook Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
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
        var email = RecipientInput.Text.Trim();
        if (string.IsNullOrEmpty(email)) return;

        if (!email.Contains('@'))
        {
            MessageBox.Show("Please enter a valid email address.", "Invalid Email",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!_recipients.Contains(email))
            _recipients.Add(email);

        RecipientInput.Text = "";
        RecipientInput.Focus();
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
            $"This will send {selectedEvents.Count * _recipients.Count} forwarding email(s).",
            "Confirm Forward", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes) return;

        try
        {
            BtnForward.IsEnabled = false;
            var app = GetOutlookApp();
            var ns = app.GetNamespace("MAPI");
            int sent = 0;
            int failed = 0;

            foreach (var ev in selectedEvents)
            {
                AppointmentItem? appt = null;
                try
                {
                    appt = (AppointmentItem)ns.GetItemFromID(ev.EntryId);
                    var forward = appt.ForwardAsVcal();
                    forward.To = string.Join(";", _recipients);
                    forward.Send();
                    sent++;
                }
                catch
                {
                    failed++;
                }
                finally
                {
                    if (appt != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appt);
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
            MessageBox.Show($"Error during forwarding:\n\n{ex.Message}",
                "Forward Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnForward.IsEnabled = true;
        }
    }
}
