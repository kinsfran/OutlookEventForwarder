using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

const string TARGET_SUBJECT = "Test meeting";
const string RECIPIENT = "karl.insfran.cs@hitachi.com";

Console.WriteLine("=== Forward Test (Actions approach) ===\n");

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

// List available actions
Console.WriteLine("\n  Available actions:");
for (int i = 1; i <= target.Actions.Count; i++)
{
    var a = target.Actions[i];
    Console.WriteLine($"    [{i}] {a.Name}");
    Marshal.ReleaseComObject(a);
}

// Forward using the built-in "Forward" action
Console.WriteLine($"\nForwarding to {RECIPIENT} via Actions[\"Forward\"]...");
try
{
    Microsoft.Office.Interop.Outlook.Action forwardAction = target.Actions["Forward"];
    Console.WriteLine($"  Action found: {forwardAction.Name}");

    object fwdObj = forwardAction.Execute();
    Console.WriteLine($"  Execute() OK! Type: {fwdObj.GetType().Name}");

    if (fwdObj is MailItem mail)
    {
        Console.WriteLine("  Got MailItem");
        mail.Recipients.Add(RECIPIENT);
        mail.Recipients.ResolveAll();
        Console.Write("  Sending... ");
        mail.Send();
        Console.WriteLine("SENT!");
        Marshal.ReleaseComObject(mail);
    }
    else if (fwdObj is MeetingItem mtg)
    {
        Console.WriteLine("  Got MeetingItem");
        mtg.Recipients.Add(RECIPIENT);
        mtg.Recipients.ResolveAll();
        Console.Write("  Sending... ");
        mtg.Send();
        Console.WriteLine("SENT!");
        Marshal.ReleaseComObject(mtg);
    }
    else
    {
        Console.WriteLine($"  Unexpected type: {fwdObj.GetType().FullName}");
        // Try dynamic as fallback
        dynamic fwd = fwdObj;
        fwd.Recipients.Add(RECIPIENT);
        fwd.Recipients.ResolveAll();
        fwd.Send();
        Console.WriteLine("  SENT (via dynamic)!");
        Marshal.ReleaseComObject(fwdObj);
    }

    Marshal.ReleaseComObject(forwardAction);
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
