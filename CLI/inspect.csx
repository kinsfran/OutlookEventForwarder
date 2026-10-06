#r "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Outlook.dll"
using System;
using System.Reflection;

var asm = Assembly.LoadFrom(@"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Outlook.dll");

// Check _AppointmentItem for Forward
var iface = asm.GetType("Microsoft.Office.Interop.Outlook._AppointmentItem");
if (iface != null)
{
    Console.WriteLine("=== _AppointmentItem members with 'Forward' ===");
    foreach (var m in iface.GetMembers())
    {
        if (m.Name.Contains("Forward", StringComparison.OrdinalIgnoreCase))
            Console.WriteLine($"  {m.MemberType}: {m.Name}");
    }
}

// Check AppointmentItem coclass
var coclass = asm.GetType("Microsoft.Office.Interop.Outlook.AppointmentItem");
if (coclass != null)
{
    Console.WriteLine("\n=== AppointmentItem members with 'Forward' ===");
    foreach (var m in coclass.GetMembers())
    {
        if (m.Name.Contains("Forward", StringComparison.OrdinalIgnoreCase))
            Console.WriteLine($"  {m.MemberType}: {m.Name}");
    }
}

// Check ItemEvents_10_Event
var events = asm.GetType("Microsoft.Office.Interop.Outlook.ItemEvents_10_Event");
if (events != null)
{
    Console.WriteLine("\n=== ItemEvents_10_Event members with 'Forward' ===");
    foreach (var m in events.GetMembers())
    {
        if (m.Name.Contains("Forward", StringComparison.OrdinalIgnoreCase))
            Console.WriteLine($"  {m.MemberType}: {m.Name}");
    }
}

// Check all types with Forward
Console.WriteLine("\n=== All types containing 'Forward' method ===");
foreach (var t in asm.GetTypes())
{
    foreach (var m in t.GetMembers(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
    {
        if (m.Name == "Forward" && m.MemberType == MemberTypes.Method)
            Console.WriteLine($"  {t.FullName}.{m.Name} ({m.MemberType})");
    }
}
