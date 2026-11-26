param(
    [Parameter(Mandatory = $true)][string]$Control,
    [Parameter(Mandatory = $true)][string]$Candidate,
    [Parameter(Mandatory = $true)][uint32]$ExpectedPages,
    [switch]$ReportOnly
)
$ErrorActionPreference = 'Stop'

# The Windows PowerShell host projects WinRT types through .NET Framework.
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
$StorageFile = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$PdfDocument = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
$PdfOptions = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType=WindowsRuntime]
$MemoryStream = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$BitmapEncoder = [Windows.Graphics.Imaging.BitmapEncoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$operations = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1
}
$asResult = $operations | Where-Object {
    $_.IsGenericMethodDefinition -and $_.GetGenericArguments().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
}
$asAction = $operations | Where-Object {
    -not $_.IsGenericMethod -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
}

function Open-Pdf([string]$Path) {
    $fileOperation = $StorageFile::GetFileFromPathAsync([IO.Path]::GetFullPath($Path))
    $fileTask = $asResult.MakeGenericMethod($StorageFile).Invoke($null, @($fileOperation))
    $file = $fileTask.GetAwaiter().GetResult()
    $loadOperation = $PdfDocument::LoadFromFileAsync($file)
    $loadTask = $asResult.MakeGenericMethod($PdfDocument).Invoke($null, @($loadOperation))
    return $loadTask.GetAwaiter().GetResult()
}

function Get-PagePaint($Document, [uint32]$Index) {
    $page = $Document.GetPage($Index)
    try {
        $width = [uint32][Math]::Round($page.Size.Width)
        $height = [uint32][Math]::Round($page.Size.Height)
        if (-not $width -or -not $height -or $width -gt 5000 -or $height -gt 5000) {
            throw 'Unexpected physical page size'
        }
        # PdfPage.Size is in DIPs. Use the corresponding 96 DPI pixel grid.
        $options = $PdfOptions::new()
        $options.DestinationWidth = $width
        $options.DestinationHeight = $height
        $options.IsIgnoringHighContrast = $true
        $options.BitmapEncoderId = $BitmapEncoder::PngEncoderId

        $stream = $MemoryStream::new()
        try {
            $renderTask = $asAction.Invoke($null, @($page.RenderToStreamAsync($stream, $options)))
            $renderTask.GetAwaiter().GetResult() | Out-Null
            $reader = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream.GetInputStreamAt(0))
            try {
                $bitmap = [Drawing.Bitmap]::new($reader)
                try {
                    if ($bitmap.Width -ne $width -or $bitmap.Height -ne $height) {
                        throw 'Rendered bitmap did not cover the physical sheet'
                    }
                    $area = [Drawing.Rectangle]::new(0, 0, $width, $height)
                    $format = [Drawing.Imaging.PixelFormat]::Format32bppArgb
                    $pixels = $bitmap.Clone($area, $format)
                    try {
                        $locked = $pixels.LockBits($area, [Drawing.Imaging.ImageLockMode]::ReadOnly, $format)
                        try {
                            if ($locked.Stride -ne $width * 4) { throw 'Unexpected bitmap row size' }
                            $bytes = [byte[]]::new($locked.Stride * $height)
                            [Runtime.InteropServices.Marshal]::Copy($locked.Scan0, $bytes, 0, $bytes.Length)
                            $hash = [Security.Cryptography.SHA256]::Create()
                            try { return "$width/${height}:" + [BitConverter]::ToString($hash.ComputeHash($bytes)) }
                            finally { $hash.Dispose() }
                        } finally { $pixels.UnlockBits($locked) }
                    } finally { $pixels.Dispose() }
                } finally { $bitmap.Dispose() }
            } finally { $reader.Dispose() }
        } finally { $stream.Dispose() }
    } finally { $page.Dispose() }
}

$left = Open-Pdf $Control
$right = Open-Pdf $Candidate
if (-not $ExpectedPages -or $left.PageCount -ne $ExpectedPages -or $right.PageCount -ne $ExpectedPages) {
    throw 'The PDF page counts differ from the pinned case'
}
$mismatches = [Collections.Generic.List[int]]::new()
for ($i = 0; $i -lt $ExpectedPages; $i++) {
    if ((Get-PagePaint $left $i) -cne (Get-PagePaint $right $i)) { $mismatches.Add($i + 1) }
}
Write-Host "WINDOWS_PDF_PIXEL_COMPARE pages $ExpectedPages exact $($ExpectedPages - $mismatches.Count) different $($mismatches.Count) dpi 96"
if ($mismatches.Count) {
    Write-Host "WINDOWS_PDF_PIXEL_PAGES $($mismatches -join ',')"
    if (-not $ReportOnly) { exit 1 }
}
