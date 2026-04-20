import EventKit
import Foundation

let store = EKEventStore()
let sema = DispatchSemaphore(value: 0)
store.requestFullAccessToEvents { _, _ in sema.signal() }
sema.wait()

let cal = Calendar.current
let now = Date()
let weekday = cal.component(.weekday, from: now)
let daysToMon = weekday == 1 ? -6 : -(weekday - 2)
let monday = cal.date(byAdding: .day, value: daysToMon, to: cal.startOfDay(for: now))!
let endOfWeek = cal.date(byAdding: .day, value: 7, to: monday)!

let skip = [
    "US Holidays", "Siri Suggestions", "Birthdays",
    "Scheduled Reminders", "Home", "iCloud", "Work", "Calendar"
]

let pred = store.predicateForEvents(withStart: monday, end: endOfWeek, calendars: nil)
var out = [[String: Any]]()
let fmt = ISO8601DateFormatter()

for ev in store.events(matching: pred) {
    let calTitle = ev.calendar?.title ?? ""
    if skip.contains(calTitle) || ev.isAllDay { continue }

    var atts = [[String: String]]()
    for a in ev.attendees ?? [] {
        let email = a.url.absoluteString.replacingOccurrences(of: "mailto:", with: "")
        atts.append(["email": email, "displayName": a.name ?? ""])
    }

    out.append([
        "id":          ev.eventIdentifier ?? "",
        "summary":     ev.title ?? "",
        "start":       ["dateTime": fmt.string(from: ev.startDate)],
        "end":         ["dateTime": fmt.string(from: ev.endDate)],
        "description": ev.notes ?? "",
        "status":      "confirmed",
        "attendees":   atts,
        "_calName":    calTitle
    ])
}

print(String(data: try! JSONSerialization.data(withJSONObject: out), encoding: .utf8)!)
