param (
    [Parameter(Mandatory)]
    [string]$ArtefactFolder,
    [Parameter(Mandatory)]
    [string]$base64
)

cd -Path "$ArtefactFolder\bin\Debug\net8.0\" -PassThru
$fileExec = "$ArtefactFolder\bin\Debug\net8.0\Coginov.GraphApi.Library.dll";
$buffer = [System.Convert]::FromBase64String($base64);
$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($buffer);
Set-AuthenticodeSignature -FilePath $fileExec -Certificate $certificate;

$certPath = "$ArtefactFolder\tempCert.pfx"
[System.IO.File]::WriteAllBytes($certPath, $buffer)

# Firmar el ensamblado usando signtool con SHA-256
Start-Process -FilePath "signtool.exe" -ArgumentList @(
    "sign",
    "/f", $certPath,
    "/fd", "SHA256",
    "/tr", "http://timestamp.digicert.com",
    "/td", "SHA256",
    $fileExec
) -Wait

Remove-Item $certPath -Force
