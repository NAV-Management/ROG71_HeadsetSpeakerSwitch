#Compress Source Script and convert to base64
$s = [System.IO.File]::ReadAllText("C:\repo\PS_Source.ps1", [System.Text.Encoding]::UTF8)
$ms = New-Object System.IO.MemoryStream
$cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
$sw = New-Object System.IO.StreamWriter($cs)
$sw.Write($s)
$sw.Close();
$sgzipb64 = [System.Convert]::ToBase64String($ms.ToArray())

$sgzipb64 | Set-Clipboard
#manually Update $s string inside batch file between '-Signs


#convert from base64 and decompress source, only used inside batch
$data = [System.Convert]::FromBase64String($sgzipb64)
$ms = New-Object System.IO.MemoryStream
$ms.Write($data, 0, $data.Length)
$ms.Seek(0,0) | Out-Null
$sr = New-Object System.IO.StreamReader(New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Decompress))
$s = $sr.ReadToEnd()