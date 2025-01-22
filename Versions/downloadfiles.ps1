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

# Crear carpeta
New-Item -ItemType Directory -Path $folderPath -Force

# Rutas locales
$localPathPKCS = "$folderPath\$SMfilePKCS"
$localPathCRT = "$folderPath\$SMfileCogCRT"
$localPathDgcCA = "$folderPath\$SMfileDgcCA"
$localPathTR = "$folderPath\$SMfileTR"

# Función para descargar archivos
function Download-File {
    param (
        [string]$storageUrl,
        [string]$sasToken,
        [string]$localPath
    )
    if (-not $storageUrl -or -not $sasToken) {
        Write-Host "Error: La URL o el token SAS están vacíos."
        return
    }
    $uri = "$storageUrl?$sasToken"
    try {
        Write-Host "Descargando archivo desde: $uri"
        Invoke-WebRequest -Uri $uri -OutFile $localPath
        Write-Host "Archivo descargado: $localPath"
    } catch {
        Write-Host "Error al descargar archivo desde: $uri. Detalles: $_"
    }
}

# Descargar archivos
Download-File -storageUrl $SMstorageUrlPKCS -sasToken $SMsasTokenPKCS -localPath $localPathPKCS
Download-File -storageUrl $SMstorageUrlCogCRT -sasToken $SMsasTokenCogCRT -localPath $localPathCRT
Download-File -storageUrl $SMstorageUrlDgcCA -sasToken $SMsasTokenDgcCA -localPath $localPathDgcCA
Download-File -storageUrl $SMstorageUrlTR -sasToken $SMsasTokenTR -localPath $localPathTR
