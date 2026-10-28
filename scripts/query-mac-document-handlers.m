#import <AppKit/AppKit.h>

// Query actual Finder Open With candidates without changing any defaults.
int main(int argc, const char *argv[]) {
    @autoreleasepool {
        NSMutableDictionary *handlers = [NSMutableDictionary dictionary];
        for (int i = 1; i < argc; i++) {
            NSString *filename = [NSString stringWithUTF8String:argv[i]];
            NSMutableArray *paths = [NSMutableArray array];
            for (NSURL *url in [[NSWorkspace sharedWorkspace] URLsForApplicationsToOpenURL:[NSURL fileURLWithPath:filename]]) {
                [paths addObject:url.URLByResolvingSymlinksInPath.path];
            }
            handlers[filename] = paths;
        }
        NSData *data = [NSJSONSerialization dataWithJSONObject:handlers options:NSJSONWritingSortedKeys error:NULL];
        if (!data) return 1;
        return fwrite(data.bytes, 1, data.length, stdout) == data.length ? 0 : 1;
    }
}
