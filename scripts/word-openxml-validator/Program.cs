using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json;

if (args.Length != 2 || !int.TryParse(args[1], out int expected)) {
    Console.Error.WriteLine("Usage: Validator <test-output-folder> <expected-DOCX-count>");
    return 2;
}
var files = Directory.EnumerateFiles(args[0], "*.docx", SearchOption.AllDirectories)
    .Where(file => Path.GetFileName(file).StartsWith("textbox-save") || Path.GetFileName(file).StartsWith("contextual-") || Path.GetFileName(file).StartsWith("list-save"))
    .Order().ToArray();
if (files.Length != expected) {
    Console.Error.WriteLine($"Expected {expected} actual exported DOCX files; found {files.Length}.");
    return 2;
}
bool failed = false;
foreach (var file in files) {
    try {
        using var document = WordprocessingDocument.Open(file, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).Select(error => new {
            id = error.Id,
            part = error.Part?.Uri.ToString(),
            path = error.Path?.XPath,
            description = error.Description
        }).ToArray();
        failed |= errors.Length != 0;
        Console.WriteLine(JsonSerializer.Serialize(new { file, errors }));
    } catch (Exception error) {
        failed = true;
        Console.WriteLine(JsonSerializer.Serialize(new { file, failure = error.Message }));
    }
}
return failed ? 1 : 0;
