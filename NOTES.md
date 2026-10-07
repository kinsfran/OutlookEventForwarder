# Outlook Event Forwarder - Technical Notes

## Goal

Forward multiple Outlook calendar events to multiple recipients at once, preserving the original meeting identity so all participants join the same Teams call.

## Key Requirement

Recipients must be added as attendees of the **original** meeting — not a copy, not a plain email. This is critical because forwarded events must link to the same Teams meeting.

## Approaches Tried

| Approach | Result | Why it failed |
|----------|--------|---------------|
| `AppointmentItem.Forward()` | Build error CS0079 | `Forward` is only an event handler in the 15.0 interop, not a method |
| `Actions["Forward"].Execute()` | .msg attachment | Recipient gets the event as an attached .msg file, not a proper forward |
| `ForwardAsVcal()` | .vcs attachment | Sends old vCalendar format, not a real forward |
| Copy + MeetingStatus + Send | New meeting created | Creates a separate meeting with a new ID — doesn't link to the original Teams meeting |
| Plain MailItem with body text | No calendar entry | Just a text email, no iCalendar data, nothing on recipient's calendar |
| Graph API (`/me/events/{id}/forward`) | Auth blocked | Hitachi tenant blocks first-party client IDs (AADSTS65002). Tried: Graph CLI (`14d82eec`), Microsoft Office (`d3590ed6`), Azure CLI (`04b07795`) — all rejected |

## Remaining Options

### 1. EWS (Exchange Web Services) with NTLM Auth
- Uses `CreateForward` method on calendar items
- Authenticates with Windows/NTLM credentials (no OAuth, no app registration)
- Not yet tested

### 2. Graph API with Admin-Registered App
- `POST /me/events/{id}/forward` does a true forward
- Requires IT to register an app in Azure AD portal and grant `Calendars.Read` permission
- Would be the cleanest solution but needs admin involvement

## Technical Details

- **Outlook interop:** GAC assemblies at `C:\Windows\assembly\GAC_MSIL\` with `EmbedInteropTypes=false` and `Private=true`
- **Calendar loading:** `Find()`/`FindNext()` pattern with `IncludeRecurrences=false`, manual date bounds check
- **WPF features working:** Date range filter, subject filter, Excel-style column filters, shift+click multi-select, double-click to open, recipient management with semicolon paste, column width persistence
