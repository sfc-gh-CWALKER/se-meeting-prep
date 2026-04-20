import EventKit
import Foundation

let store = EKEventStore()
let sema = DispatchSemaphore(value: 0)
store.requestFullAccessToEvents { _, _ in sema.signal() }
sema.wait()

guard CommandLine.arguments.count >= 2 else {
    fputs("Usage: update_cal_event <eventIdentifier>\n", stderr)
    exit(1)
}

let eventId  = CommandLine.arguments[1]
let newDesc  = String(data: FileHandle.standardInput.readDataToEndOfFile(), encoding: .utf8) ?? ""

guard let ev = store.event(withIdentifier: eventId) else {
    fputs("not-found\n", stderr)
    exit(1)
}

ev.notes = newDesc
try! store.save(ev, span: .thisEvent)
print("updated")
