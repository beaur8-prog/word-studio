import AppKit
import Foundation

// Query the actual Finder "Open With" candidates without changing defaults.
var handlers: [String: [String]] = [:]
for filename in CommandLine.arguments.dropFirst() {
    handlers[filename] = NSWorkspace.shared.urlsForApplications(toOpen: URL(fileURLWithPath: filename))
        .map { $0.resolvingSymlinksInPath().path }
}
let data = try JSONSerialization.data(withJSONObject: handlers, options: [.sortedKeys])
FileHandle.standardOutput.write(data)
