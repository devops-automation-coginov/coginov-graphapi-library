param (
    [Parameter(Mandatory)]
    [string]$SMstorageUrlPKCS,
    [Parameter(Mandatory)]
    [string]$SMsasTokenPKCS,
    [Parameter(Mandatory)]
	[string]$SMfilePKCS,
	
    [Parameter(Mandatory)]
    [string]$SMstorageUrlCogCRT,
    [Parameter(Mandatory)]
    [string]$SMsasTokenCogCRT,
	[Parameter(Mandatory)]
	[string]$SMfileCogCRT,
	
    [Parameter(Mandatory)]
    [string]$SMstorageUrlDgcCA,
    [Parameter(Mandatory)]
    [string]$SMsasTokenDgcCA,
	[Parameter(Mandatory)]
	[string]$SMfileDgcCA,
	
	[Parameter(Mandatory)]
    [string]$SMstorageUrlTR,
    [Parameter(Mandatory)]
    [string]$SMsasTokenTR,
	[Parameter(Mandatory)]
	[string]$SMfileTR
)
$folderPath = "C:\signStuff"
$localPathPKCS = "$folderPath\$SMfilePKCS"
$localPathCRT = "$folderPath\SMfileCogCRT"
$localPathDgcCA = "$folderPath\$SMfileDgcCA"
$localPathTR = "$folderPath\$SMfileTR"
New-Item -ItemType Directory -Path $folderPath -Force

# Downloading file
Invoke-WebRequest -Uri "$storageUrlPKCS?$sasTokenPKCS" -OutFile $localPathPKCS
Write-Host "Archivo descargado: $localPathPKCS"

# Downloading file
Invoke-WebRequest -Uri "$storageUrlCogCRT?$sasTokenCogCRT" -OutFile $localPathCRT
Write-Host "Archivo descargado: $localPathCRT"
	  
# Downloading file
Invoke-WebRequest -Uri "$storageUrlDgcCA?$sasTokenDgcCA" -OutFile $localPathDgcCA
Write-Host "Archivo descargado: $localPathDgcCA"

# Downloading file
Invoke-WebRequest -Uri "$storageUrlTR?$sasTokenTR" -OutFile $localPathTR
Write-Host "Archivo descargado: $localPathTR"

dir "$folderPath"