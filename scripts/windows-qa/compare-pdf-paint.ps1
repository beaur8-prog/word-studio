param(
    [Parameter(Mandatory = $true)][string]$Control,
    [Parameter(Mandatory = $true)][string]$Candidate,
    [Parameter(Mandatory = $true)][uint32]$ExpectedPages,
    [switch]$ReportOnly,
    [switch]$MeasureDelta
)
$ErrorActionPreference = 'Stop'

# The Windows PowerShell host projects WinRT types through .NET Framework.
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
if ($MeasureDelta -and -not ('PixelDelta' -as [type])) {
    Add-Type -TypeDefinition @'
public static class PixelDelta {
    public static int[] Compare(byte[] a, byte[] b, int width) {
        int x0 = width, y0 = a.Length / (4 * width), x1 = -1, y1 = -1, count = 0, max = 0;
        long absolute = 0;
        for (int i = 0; i < a.Length; i += 4) {
            int difference = System.Math.Max(System.Math.Abs(a[i] - b[i]),
                System.Math.Max(System.Math.Abs(a[i+1] - b[i+1]), System.Math.Abs(a[i+2] - b[i+2])));
            if (difference == 0) continue;
            int x = (i / 4) % width;
            int y = (i / 4) / width;
            x0 = System.Math.Min(x0, x);
            y0 = System.Math.Min(y0, y);
            x1 = System.Math.Max(x1, x);
            y1 = System.Math.Max(y1, y);
            count++;
            max = System.Math.Max(max, difference);
            absolute += difference;
        }
        int mean = count == 0 ? 0 : (int)(100 * absolute / count);
        return new int[] { count, x0, y0, x1, y1, max, mean };
    }
}
'@
}
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
                            try {
                                $key = "$width/${height}:" + [BitConverter]::ToString($hash.ComputeHash($bytes))
                                if ($MeasureDelta) { return @{ Key=$key; Width=$width; Height=$height; Pixels=$bytes } }
                                return $key
                            }
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
    $a = Get-PagePaint $left $i
    $b = Get-PagePaint $right $i
    if ($MeasureDelta) {
        if ($a.Key -ceq $b.Key) { continue }
        $mismatches.Add($i + 1)
        if ($a.Width -ne $b.Width -or $a.Height -ne $b.Height) { throw 'Cannot compare different physical raster sizes' }
        $delta = [PixelDelta]::Compare($a.Pixels, $b.Pixels, $a.Width)
        Write-Host "WINDOWS_PDF_DELTA $($i + 1) $($a.Width)x$($a.Height) changed=$($delta[0]) bbox=$($delta[1]),$($delta[2]),$($delta[3]),$($delta[4]) max=$($delta[5]) meanHundredths=$($delta[6])"
    } elseif ($a -cne $b) {
        $mismatches.Add($i + 1)
    }
}
Write-Host "WINDOWS_PDF_PIXEL_COMPARE pages $ExpectedPages exact $($ExpectedPages - $mismatches.Count) different $($mismatches.Count) dpi 96"
if ($mismatches.Count) {
    Write-Host "WINDOWS_PDF_PIXEL_PAGES $($mismatches -join ',')"
    if (-not $ReportOnly) { exit 1 }
}
