param (
    [Parameter(Mandatory)]
    [string]$ArtefactFolder,
    [Parameter(Mandatory)]
    [string]$base64
)

cd -Path "$ArtefactFolder\Coginov.Semantic.Library\bin\Debug\netstandard2.1\" -PassThru
$fileExec = "$ArtefactFolder\Coginov.Semantic.Library\bin\Debug\netstandard2.1\Coginov.Semantic.Library.dll";
$buffer = [System.Convert]::FromBase64String($base64);
$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($buffer);
Set-AuthenticodeSignature -FilePath $fileExec -Certificate $certificate;