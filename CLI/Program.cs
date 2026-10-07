using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

const string TARGET_SUBJECT = "Test meeting";
const string RECIPIENT = "karl.insfran.cs@hitachi.com";

Console.WriteLine("=== Forward Test (manual MailItem) ===\n");

var app = new OutlookApp();
var ns = app.GetNamespace("MAPI");

var calendarFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
var items = calendarFolder.Items;
items.Sort("[Start]", false);
items.IncludeRecurrences = false;

var from = DateTime.Today.ToString("g");
var to = DateTime.Today.AddDays(7).ToString("g");
var filter = $"[Start] >= '{from}' AND [End] <= '{to}'";

Console.Write("Finding 'Test meeting'... ");
AppointmentItem? target = null;
object? found = items.Find(filter);
while (found != null)
{
    if (found is AppointmentItem appt && appt.Subject == TARGET_SUBJECT)
    {
        target = appt;
        Console.WriteLine($"Found! Start: {appt.Start:yyyy-MM-dd HH:mm}");
        break;
    }
    if (found is not null) Marshal.ReleaseComObject(found);
    found = items.FindNext();
}

if (target == null)
{
    Console.WriteLine("NOT FOUND. Exiting.");
    return;
}

Console.WriteLine($"  IsRecurring: {target.IsRecurring}");
Console.WriteLine($"  MeetingStatus: {target.MeetingStatus}");
Console.WriteLine($"  Location: {target.Location}");
Console.WriteLine($"  Organizer: {target.Organizer}");

Console.WriteLine($"\nForwarding to {RECIPIENT} (manual MailItem)...");
try
{
    var mail = (MailItem)app.CreateItem(OlItemType.olMailItem);
    mail.Subject = "FW: " + target.Subject;

    var body = new StringBuilder();
    body.AppendLine("---------- Forwarded event ----------");
    body.AppendLine($"Subject: {target.Subject}");
    body.AppendLine($"When: {target.Start:dddd, MMMM d, yyyy h:mm tt} - {target.End:h:mm tt}");
    if (!string.IsNullOrEmpty(target.Location))
        body.AppendLine($"Location: {target.Location}");
    body.AppendLine($"Organizer: {target.Organizer}");
    if (!string.IsNullOrEmpty(target.Body))
    {
        body.AppendLine();
        body.AppendLine(target.Body);
    }

    mail.Body = body.ToString();
    mail.Recipients.Add(RECIPIENT);
    mail.Recipients.ResolveAll();

    Console.WriteLine($"  Subject: {mail.Subject}");
    Console.WriteLine($"  Body preview: {mail.Body[..Math.Min(200, mail.Body.Length)]}");
    Console.Write("  Sending... ");
    mail.Send();
    Console.WriteLine("SENT!");
    Marshal.ReleaseComObject(mail);
}
catch (System.Exception ex)
{
    Console.WriteLine($"  FAILED: {ex.GetType().Name}: {ex.Message}");
    if (ex.InnerException != null)
        Console.WriteLine($"  Inner: {ex.InnerException.Message}");
}

Console.WriteLine("\n--- Done ---");
Marshal.ReleaseComObject(target);
Marshal.ReleaseComObject(items);
Marshal.ReleaseComObject(calendarFolder);
Marshal.ReleaseComObject(ns);
