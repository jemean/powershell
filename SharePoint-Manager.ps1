#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint File Management via Microsoft Graph API (sans module additionnel)

.DESCRIPTION
    Gestion professionnelle de fichiers SharePoint via l'API Microsoft Graph.
    Authentification par certificat (flux client_credentials + JWT assertion).

.PARAMETER Action
    UploadFile | UploadDirectory | DownloadFile | DownloadDirectory |
    ListFiles | DeleteFile | DeleteByPattern | TestFileExists

.PARAMETER LocalPath        Chemin du fichier local (upload/download unitaire)
.PARAMETER LocalDirectory   Répertoire local (upload/download en lot)
.PARAMETER SPFolder         Chemin du dossier SharePoint (ex: "Documents/Rapports")
.PARAMETER FileName         Nom du fichier dans SharePoint
.PARAMETER Pattern          Filtre wildcard (ex: "*.csv", "rapport_*.xlsx")  [défaut: *]
.PARAMETER DestinationPath  Répertoire de destination local (download)
.PARAMETER PostProcess      None | Delete | Archive  — traitement du source après succès
.PARAMETER ArchivePath      Répertoire local d'archivage (si PostProcess = Archive)
.PARAMETER ConfigFile       Chemin du fichier de configuration JSON
                            [défaut: SharePoint-Config.json dans le répertoire du script]

.EXAMPLE
    .\SharePoint-Manager.ps1 -Action UploadFile -LocalPath "C:\data\rapport.xlsx" -SPFolder "Documents/Rapports"

.EXAMPLE
    .\SharePoint-Manager.ps1 -Action UploadDirectory -LocalDirectory "C:\exports" -Pattern "*.csv" `
        -SPFolder "Data/Imports" -PostProcess Archive -ArchivePath "C:\exports\done"

.EXAMPLE
    .\SharePoint-Manager.ps1 -Action DownloadDirectory -SPFolder "Data/Exports" -Pattern "*.xlsx" `
        -DestinationPath "C:\downloads" -PostProcess Delete

.EXAMPLE
    .\SharePoint-Manager.ps1 -Action ListFiles -SPFolder "Documents/Rapports" -Pattern "2024_*"

.NOTES
    Prérequis :
      - Enregistrement d'application Azure AD avec permission Sites.ReadWrite.All (application)
      - Certificat chargé dans l'enregistrement d'application (onglet Certificats & secrets)
      - Certificat (avec clé privée) installé dans CurrentUser\My ou LocalMachine\My
      - Renseigner la section #region Configuration ci-dessous
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [ValidateSet('UploadFile','UploadDirectory','DownloadFile','DownloadDirectory',
                 'ListFiles','DeleteFile','DeleteByPattern','TestFileExists')]
    [string]$Action,

    [string]$LocalPath,
    [string]$LocalDirectory,
    [string]$SPFolder        = '',
    [string]$FileName        = '',
    [string]$Pattern         = '*',
    [string]$DestinationPath,

    [ValidateSet('None','Delete','Archive')]
    [string]$PostProcess = 'None',
    [string]$ArchivePath,

    [string]$ConfigFile = (Join-Path $PSScriptRoot 'SharePoint-Config.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
#region Configuration
# ============================================================
$Script:Config        = $null   # initialisé par Import-SPConfig
$Script:TokenCache    = @{ AccessToken = $null; ExpiresAt = [DateTime]::MinValue }
$Script:CachedSiteId  = $null
$Script:CachedDriveId = $null

function Import-SPConfig {
    <#
    .SYNOPSIS Charge et valide le fichier de configuration JSON.
    Appelé automatiquement à l'exécution, ou manuellement après dot-sourcing :
        . .\SharePoint-Manager.ps1
        Import-SPConfig -ConfigFile "C:\projets\projet1\config.json"
    #>
    param(
        [Parameter(Mandatory)][string]$ConfigFile
    )

    if (-not (Test-Path $ConfigFile -PathType Leaf)) {
        throw "Fichier de configuration introuvable : $ConfigFile`nCréer un fichier JSON à partir de SharePoint-Config.sample.json"
    }

    $json = Get-Content $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json

    # Validation des champs obligatoires
    foreach ($field in 'TenantId','ClientId','CertThumbprint','SharePointHost','SitePath') {
        $val = $json.$field
        if ([string]::IsNullOrWhiteSpace($val) -or $val -match '^[xX]+(-[xX]+)*$') {
            throw "Champ '$field' manquant ou non renseigné dans '$ConfigFile'."
        }
    }

    $dir = Split-Path $ConfigFile -Parent

    $Script:Config = [ordered]@{
        # Obligatoires — issus du fichier de config
        TenantId           = $json.TenantId
        ClientId           = $json.ClientId
        CertThumbprint     = $json.CertThumbprint
        SharePointHost     = $json.SharePointHost
        SitePath           = $json.SitePath

        # Optionnels — valeur du fichier ou défaut
        DriveName          = if ($null -ne $json.DriveName)          { $json.DriveName }                           else { 'Documents' }
        LargeFileThreshold = if ($null -ne $json.LargeFileThresholdMB) { [long]$json.LargeFileThresholdMB * 1MB } else { 4MB }
        UploadChunkSize    = if ($null -ne $json.UploadChunkSizeMB)  { [long]$json.UploadChunkSizeMB * 1MB }       else { 10485760 }
        MaxRetries         = if ($null -ne $json.MaxRetries)         { [int]$json.MaxRetries }                     else { 3 }
        RetryDelaySeconds  = if ($null -ne $json.RetryDelaySeconds)  { [int]$json.RetryDelaySeconds }              else { 5 }
        LogFile            = if ($json.LogFile)                      { $json.LogFile }                             else { Join-Path $dir 'SharePoint-Manager.log' }
        LogLevel           = if ($json.LogLevel)                     { $json.LogLevel }                            else { 'INFO' }
        LogToConsole       = if ($null -ne $json.LogToConsole)       { [bool]$json.LogToConsole }                  else { $true }

        # Fixes — non exposés dans le fichier de config
        GraphBaseUrl       = 'https://graph.microsoft.com/v1.0'
        TokenEndpoint      = 'https://login.microsoftonline.com'
        Scope              = 'https://graph.microsoft.com/.default'
    }

    # Réinitialisation des caches à chaque chargement de config
    $Script:TokenCache    = @{ AccessToken = $null; ExpiresAt = [DateTime]::MinValue }
    $Script:CachedSiteId  = $null
    $Script:CachedDriveId = $null

    Write-Verbose "Config chargée : $($Script:Config.SharePointHost)$($Script:Config.SitePath)"
}
#endregion

# ============================================================
#region Journalisation
# ============================================================
function Write-Log {
    [CmdletBinding()]
    param(
        [ValidateSet('DEBUG','INFO','WARNING','ERROR')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory)][string]$Message,
        [System.Exception]$Exception
    )

    $levels = @{ DEBUG = 0; INFO = 1; WARNING = 2; ERROR = 3 }
    if ($levels[$Level] -lt $levels[$Script:Config.LogLevel]) { return }

    $ts       = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry    = "[$ts] [$Level] $Message"
    if ($Exception) { $entry += "`n  >> $($Exception.Message)" }

    try { $entry | Out-File -FilePath $Script:Config.LogFile -Append -Encoding UTF8 } catch {}

    if ($Script:Config.LogToConsole) {
        switch ($Level) {
            'DEBUG'   { Write-Verbose $entry }
            'INFO'    { Write-Host    $entry -ForegroundColor Cyan }
            'WARNING' { Write-Warning $Message }
            'ERROR'   { Write-Host    $entry -ForegroundColor Red }
        }
    }
}
#endregion

# ============================================================
#region Authentification par certificat
# ============================================================
function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes) -replace '\+','-' -replace '/','_' -replace '=+$',''
}

function New-JwtAssertion {
    <#
    .SYNOPSIS Construit et signe le JWT client_assertion pour le flux certificate credentials.
    #>
    [OutputType([string])]
    param()

    $thumbprint = $Script:Config.CertThumbprint -replace '\s',''

    $cert = Get-Item "Cert:\CurrentUser\My\$thumbprint"  -ErrorAction SilentlyContinue
    if (-not $cert) {
        $cert = Get-Item "Cert:\LocalMachine\My\$thumbprint" -ErrorAction SilentlyContinue
    }
    if (-not $cert)             { throw "Certificat introuvable (thumbprint: $thumbprint). Vérifiez CurrentUser\My et LocalMachine\My." }
    if (-not $cert.HasPrivateKey) { throw "Le certificat $thumbprint n'a pas de clé privée accessible." }

    Write-Log -Level DEBUG -Message "Certificat utilisé : $($cert.Subject)"

    # x5t : thumbprint hex → bytes → Base64Url
    $thumbBytes = [byte[]]$(for ($i = 0; $i -lt $thumbprint.Length; $i += 2) {
        [Convert]::ToByte($thumbprint.Substring($i, 2), 16)
    })

    $headerJson = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = (ConvertTo-Base64Url $thumbBytes) } |
                  ConvertTo-Json -Compress
    $headerEnc  = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($headerJson))

    # Epoch compatible PS 5.1 et PS 7
    $epoch = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
    $now   = [long]([DateTime]::UtcNow - $epoch).TotalSeconds

    $payloadJson = [ordered]@{
        aud = "$($Script:Config.TokenEndpoint)/$($Script:Config.TenantId)/oauth2/v2.0/token"
        iss = $Script:Config.ClientId
        sub = $Script:Config.ClientId
        jti = [Guid]::NewGuid().ToString()
        nbf = $now
        exp = $now + 600
    } | ConvertTo-Json -Compress
    $payloadEnc = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($payloadJson))

    $signingInput = "$headerEnc.$payloadEnc"

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if (-not $rsa) { throw "Impossible d'accéder à la clé privée RSA du certificat." }

    $sigBytes = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($signingInput),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    return "$signingInput.$(ConvertTo-Base64Url $sigBytes)"
}

function Get-GraphToken {
    <#
    .SYNOPSIS Retourne un token Graph API valide (cache de 60s avant expiration).
    #>
    [OutputType([string])]
    param()

    if ($Script:TokenCache.AccessToken -and
        [DateTime]::UtcNow -lt $Script:TokenCache.ExpiresAt.AddSeconds(-60)) {
        Write-Log -Level DEBUG -Message 'Token Graph API depuis le cache.'
        return $Script:TokenCache.AccessToken
    }

    Write-Log -Level INFO -Message 'Acquisition d''un nouveau token Graph API...'

    $endpoint = "$($Script:Config.TokenEndpoint)/$($Script:Config.TenantId)/oauth2/v2.0/token"
    $body = @{
        grant_type            = 'client_credentials'
        client_id             = $Script:Config.ClientId
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = New-JwtAssertion
        scope                 = $Script:Config.Scope
    }

    $resp = Invoke-RestMethod -Uri $endpoint -Method POST -Body $body `
        -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop

    $Script:TokenCache.AccessToken = $resp.access_token
    $Script:TokenCache.ExpiresAt   = [DateTime]::UtcNow.AddSeconds($resp.expires_in)
    Write-Log -Level INFO -Message "Token obtenu. Expire à $($Script:TokenCache.ExpiresAt.ToString('HH:mm:ss')) UTC."
    return $Script:TokenCache.AccessToken
}
#endregion

# ============================================================
#region Graph API — wrapper HTTP avec retry
# ============================================================
function Invoke-GraphRequest {
    <#
    .SYNOPSIS Appel générique à l'API Graph avec gestion du retry (429 / 5xx).

    .PARAMETER Endpoint    Chemin relatif ou URL complète
    .PARAMETER Method      GET | POST | PUT | PATCH | DELETE
    .PARAMETER Body        Hashtable/PSObject (sérialisé JSON) ou byte[] (binaire)
    .PARAMETER ContentType Content-Type (défaut : application/json)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Endpoint,
        [ValidateSet('GET','POST','PUT','PATCH','DELETE')]
        [string]$Method      = 'GET',
        [object]$Body        = $null,
        [string]$ContentType = 'application/json'
    )

    $uri     = if ($Endpoint -match '^https?://') { $Endpoint } `
               else { "$($Script:Config.GraphBaseUrl)/$Endpoint" }
    $attempt = 0

    do {
        $attempt++
        $token   = Get-GraphToken
        $headers = @{
            Authorization = "Bearer $token"
            Accept        = 'application/json'
        }

        $params = @{ Uri = $uri; Method = $Method; Headers = $headers; ErrorAction = 'Stop' }

        if ($Body) {
            if ($Body -is [byte[]]) {
                $params.Body        = $Body
                $params.ContentType = $ContentType
            } else {
                $params.Body        = [Text.Encoding]::UTF8.GetBytes(
                    $(if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 -Compress })
                )
                $params.ContentType = $ContentType
            }
        }

        try {
            return Invoke-RestMethod @params
        } catch {
            # Récupération du code HTTP (PS 5.1 et PS 7+)
            $statusCode = 0
            try {
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
            } catch {}

            if ($statusCode -eq 404)   { throw }  # pas de retry sur 404
            if ($statusCode -eq 401)   { $Script:TokenCache.AccessToken = $null }  # forcer renouvellement

            $retryable = $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600) -or $statusCode -eq 0
            if (-not $retryable -or $attempt -ge $Script:Config.MaxRetries) { throw }

            $delay = $Script:Config.RetryDelaySeconds
            try { if ($_.Exception.Response.Headers['Retry-After']) { $delay = [int]$_.Exception.Response.Headers['Retry-After'] } } catch {}

            Write-Log -Level WARNING -Message "HTTP $statusCode — retry $attempt/$($Script:Config.MaxRetries) dans ${delay}s..."
            Start-Sleep -Seconds $delay
        }
    } while ($true)
}
#endregion

# ============================================================
#region Helpers SharePoint
# ============================================================
function Get-SPSiteId {
    [OutputType([string])] param()
    if ($Script:CachedSiteId) { return $Script:CachedSiteId }

    $h    = $Script:Config.SharePointHost
    $p    = $Script:Config.SitePath.TrimStart('/')
    $site = Invoke-GraphRequest -Endpoint "sites/${h}:/${p}"
    $Script:CachedSiteId = $site.id
    Write-Log -Level DEBUG -Message "Site ID : $($Script:CachedSiteId)"
    return $Script:CachedSiteId
}

function Get-SPDriveId {
    [OutputType([string])] param()
    if ($Script:CachedDriveId) { return $Script:CachedDriveId }

    $siteId    = Get-SPSiteId
    $driveName = $Script:Config.DriveName

    if ([string]::IsNullOrWhiteSpace($driveName)) {
        $Script:CachedDriveId = (Invoke-GraphRequest -Endpoint "sites/$siteId/drive").id
    } else {
        $drives = Invoke-GraphRequest -Endpoint "sites/$siteId/drives"
        $drive  = $drives.value | Where-Object { $_.name -eq $driveName }
        if (-not $drive) {
            $avail = ($drives.value.name) -join ', '
            throw "Bibliothèque '$driveName' introuvable. Disponibles : $avail"
        }
        $Script:CachedDriveId = $drive.id
    }

    Write-Log -Level DEBUG -Message "Drive ID : $($Script:CachedDriveId)"
    return $Script:CachedDriveId
}

function ConvertTo-SPApiPath {
    <#
    .SYNOPSIS Encode un chemin SharePoint pour l'API Graph.
    Retourne "root" (racine) ou "root:/seg1/seg2:" (chemin encodé).
    #>
    param([string]$Path)
    $norm = $Path.Trim('/','\') -replace '\\','/'
    if ([string]::IsNullOrWhiteSpace($norm)) { return 'root' }
    $enc = ($norm -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
    return "root:/$enc:"
}
#endregion

# ============================================================
#region Post-processing local
# ============================================================
function Invoke-LocalPostProcess {
    <#
    .SYNOPSIS Supprime ou archive un fichier local après traitement réussi.
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('None','Delete','Archive')][string]$Action = 'None',
        [string]$ArchivePath
    )
    switch ($Action) {
        'None'   { return }
        'Delete' {
            Write-Log -Level INFO -Message "Post-process : suppression locale '$FilePath'"
            Remove-Item -Path $FilePath -Force
        }
        'Archive' {
            if (-not $ArchivePath) { throw "-ArchivePath requis avec PostProcess=Archive" }
            if (-not (Test-Path $ArchivePath)) { New-Item -Path $ArchivePath -ItemType Directory -Force | Out-Null }

            $dest = Join-Path $ArchivePath (Split-Path $FilePath -Leaf)
            if (Test-Path $dest) {
                $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
                $base = [IO.Path]::GetFileNameWithoutExtension($dest)
                $ext  = [IO.Path]::GetExtension($dest)
                $dest = Join-Path $ArchivePath "${base}_${ts}${ext}"
            }
            Write-Log -Level INFO -Message "Post-process : archivage local '$dest'"
            Move-Item -Path $FilePath -Destination $dest -Force
        }
    }
}
#endregion

# ============================================================
#region Upload
# ============================================================
function Invoke-SPLargeFileUpload {
    <#
    .SYNOPSIS Upload par session (fichiers > LargeFileThreshold). Usage interne.
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$SPFullPath,
        [Parameter(Mandatory)][string]$DriveId
    )

    $fi        = Get-Item $FilePath
    $apiPath   = ConvertTo-SPApiPath -Path $SPFullPath
    $chunkSize = $Script:Config.UploadChunkSize

    Write-Log -Level DEBUG -Message "Création session upload pour '$($fi.Name)' ($([Math]::Round($fi.Length/1MB,2)) Mo)"

    $session   = Invoke-GraphRequest -Method POST `
        -Endpoint "drives/$DriveId/$apiPath/createUploadSession" `
        -Body @{ item = @{ '@microsoft.graph.conflictBehavior' = 'replace' } }
    $uploadUrl = $session.uploadUrl
    $total     = $fi.Length
    $offset    = 0
    $result    = $null
    $stream    = [IO.File]::OpenRead($FilePath)
    $buffer    = New-Object byte[] $chunkSize

    try {
        while ($offset -lt $total) {
            $read = $stream.Read($buffer, 0, $chunkSize)
            if ($read -eq 0) { break }
            $chunk   = if ($read -eq $chunkSize) { $buffer } else { $buffer[0..($read - 1)] }
            $end     = $offset + $read - 1
            $pct     = [Math]::Round(($offset / $total) * 100, 1)
            Write-Progress -Activity "Upload $($fi.Name)" -Status "$pct%" -PercentComplete $pct

            $result = Invoke-RestMethod -Uri $uploadUrl -Method PUT -Body $chunk `
                -Headers @{ 'Content-Range' = "bytes $offset-$end/$total"; 'Content-Length' = "$read" } `
                -ContentType 'application/octet-stream' -ErrorAction Stop
            $offset += $read
        }
    } finally {
        $stream.Close()
        Write-Progress -Activity "Upload $($fi.Name)" -Completed
    }
    return $result
}

function Invoke-SPUploadFile {
    <#
    .SYNOPSIS Upload d'un fichier local vers un dossier SharePoint.

    .PARAMETER LocalFilePath   Fichier source local
    .PARAMETER SPFolderPath    Dossier cible SharePoint
    .PARAMETER RemoteFileName  Nom dans SharePoint (défaut : nom du fichier local)
    .PARAMETER PostProcess     None | Delete | Archive (sur le fichier local)
    .PARAMETER ArchivePath     Dossier d'archivage local
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$SPFolderPath,
        [string]$RemoteFileName,
        [ValidateSet('None','Delete','Archive')][string]$PostProcess = 'None',
        [string]$ArchivePath
    )

    if (-not (Test-Path $LocalFilePath -PathType Leaf)) { throw "Fichier introuvable : $LocalFilePath" }

    $fi         = Get-Item $LocalFilePath
    $remoteName = if ($RemoteFileName) { $RemoteFileName } else { $fi.Name }
    $spPath     = "$($SPFolderPath.TrimEnd('/'))/$remoteName"
    $driveId    = Get-SPDriveId

    if ($PSCmdlet.ShouldProcess($spPath, 'Upload SharePoint')) {
        Write-Log -Level INFO -Message "Upload : '$($fi.FullName)' → '$spPath' ($([Math]::Round($fi.Length/1KB,1)) Ko)"

        $result = if ($fi.Length -gt $Script:Config.LargeFileThreshold) {
            Invoke-SPLargeFileUpload -FilePath $fi.FullName -SPFullPath $spPath -DriveId $driveId
        } else {
            $apiPath = ConvertTo-SPApiPath -Path $spPath
            Invoke-GraphRequest -Method PUT -Endpoint "drives/$driveId/$apiPath/content" `
                -Body ([IO.File]::ReadAllBytes($fi.FullName)) -ContentType 'application/octet-stream'
        }

        Write-Log -Level INFO -Message "Upload réussi : $($result.name) (id: $($result.id))"
        Invoke-LocalPostProcess -FilePath $fi.FullName -Action $PostProcess -ArchivePath $ArchivePath
        return $result
    }
}

function Invoke-SPUploadDirectory {
    <#
    .SYNOPSIS Upload des fichiers d'un répertoire local correspondant à un pattern.

    .PARAMETER LocalDirectory  Répertoire source
    .PARAMETER Pattern         Filtre wildcard (défaut : *)
    .PARAMETER SPFolderPath    Dossier cible SharePoint
    .PARAMETER Recurse         Traiter les sous-répertoires
    .PARAMETER PostProcess     None | Delete | Archive (sur chaque fichier local)
    .PARAMETER ArchivePath     Dossier d'archivage local
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$LocalDirectory,
        [string]$Pattern = '*',
        [Parameter(Mandatory)][string]$SPFolderPath,
        [switch]$Recurse,
        [ValidateSet('None','Delete','Archive')][string]$PostProcess = 'None',
        [string]$ArchivePath
    )

    if (-not (Test-Path $LocalDirectory -PathType Container)) { throw "Répertoire introuvable : $LocalDirectory" }

    $getParams = @{ Path = $LocalDirectory; Filter = $Pattern; File = $true }
    if ($Recurse) { $getParams.Recurse = $true }
    $files = @(Get-ChildItem @getParams)

    if (-not $files) {
        Write-Log -Level WARNING -Message "Aucun fichier '$Pattern' dans '$LocalDirectory'."
        return [PSCustomObject]@{ Success = 0; Failure = 0; Errors = @() }
    }

    Write-Log -Level INFO -Message "$($files.Count) fichier(s) à uploader (pattern: $Pattern)"
    $ok = 0; $ko = 0; $errors = [System.Collections.Generic.List[object]]::new()

    foreach ($f in $files) {
        try {
            $spTarget = if ($Recurse -and $f.DirectoryName -ne (Resolve-Path $LocalDirectory).Path) {
                $rel = $f.DirectoryName.Substring($LocalDirectory.TrimEnd('\','/').Length).TrimStart('\','/')
                "$($SPFolderPath.TrimEnd('/'))/$($rel -replace '\\','/')"
            } else { $SPFolderPath }

            Invoke-SPUploadFile -LocalFilePath $f.FullName -SPFolderPath $spTarget `
                -PostProcess $PostProcess -ArchivePath $ArchivePath
            $ok++
        } catch {
            $ko++
            $errors.Add([PSCustomObject]@{ File = $f.Name; Error = $_.Exception.Message })
            Write-Log -Level ERROR -Message "Échec upload '$($f.Name)' : $_" -Exception $_.Exception
        }
    }

    Write-Log -Level INFO -Message "Upload terminé : $ok succès, $ko échec(s)"
    return [PSCustomObject]@{ Success = $ok; Failure = $ko; Errors = $errors.ToArray() }
}
#endregion

# ============================================================
#region Download
# ============================================================
function Invoke-SPDownloadFile {
    <#
    .SYNOPSIS Télécharge un fichier depuis SharePoint vers un répertoire local.

    .PARAMETER SPFolderPath    Dossier SharePoint source
    .PARAMETER FileName        Nom du fichier
    .PARAMETER DestinationPath Répertoire local de destination
    .PARAMETER PostProcess     None | Delete | Archive (sur le fichier SharePoint)
    .PARAMETER SPArchivePath   Dossier SharePoint d'archivage (si PostProcess = Archive)
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SPFolderPath,
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string]$DestinationPath,
        [ValidateSet('None','Delete','Archive')][string]$PostProcess = 'None',
        [string]$SPArchivePath
    )

    if (-not (Test-Path $DestinationPath)) { New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null }

    $spPath  = "$($SPFolderPath.TrimEnd('/'))/$FileName"
    $apiPath = ConvertTo-SPApiPath -Path $spPath
    $driveId = Get-SPDriveId
    $outFile = Join-Path $DestinationPath $FileName

    if ($PSCmdlet.ShouldProcess($spPath, 'Download SharePoint')) {
        Write-Log -Level INFO -Message "Download : '$spPath' → '$outFile'"

        # Récupération de l'URL de téléchargement pré-signée (évite les problèmes de redirect)
        $item = Invoke-GraphRequest -Endpoint "drives/$driveId/$apiPath"
        $dlUrl = $item.'@microsoft.graph.downloadUrl'
        if (-not $dlUrl) { throw "URL de téléchargement introuvable pour '$spPath'." }

        Invoke-WebRequest -Uri $dlUrl -OutFile $outFile -UseBasicParsing -ErrorAction Stop
        Write-Log -Level INFO -Message "Download réussi : '$outFile'"

        switch ($PostProcess) {
            'Delete' {
                Write-Log -Level INFO -Message "Post-process : suppression SharePoint '$spPath'"
                Remove-SPFile -SPFolderPath $SPFolderPath -FileName $FileName
            }
            'Archive' {
                if (-not $SPArchivePath) { throw "-SPArchivePath requis avec PostProcess=Archive" }
                $destFolder = Invoke-GraphRequest -Endpoint "drives/$driveId/$(ConvertTo-SPApiPath $SPArchivePath)" `
                    -ErrorAction SilentlyContinue
                if (-not $destFolder) { throw "Dossier d'archivage SharePoint introuvable : $SPArchivePath" }
                Invoke-GraphRequest -Method PATCH `
                    -Endpoint "drives/$driveId/items/$($item.id)" `
                    -Body @{ parentReference = @{ id = $destFolder.id }; name = $FileName } | Out-Null
                Write-Log -Level INFO -Message "Fichier SharePoint archivé dans '$SPArchivePath'"
            }
        }

        return $outFile
    }
}

function Invoke-SPDownloadDirectory {
    <#
    .SYNOPSIS Télécharge les fichiers d'un dossier SharePoint correspondant à un pattern.

    .PARAMETER SPFolderPath    Dossier SharePoint source
    .PARAMETER Pattern         Filtre wildcard (défaut : *)
    .PARAMETER DestinationPath Répertoire local de destination
    .PARAMETER PostProcess     None | Delete | Archive (sur les fichiers SharePoint)
    .PARAMETER SPArchivePath   Dossier SharePoint d'archivage
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SPFolderPath,
        [string]$Pattern = '*',
        [Parameter(Mandatory)][string]$DestinationPath,
        [ValidateSet('None','Delete','Archive')][string]$PostProcess = 'None',
        [string]$SPArchivePath
    )

    $files = @(Get-SPFileList -SPFolderPath $SPFolderPath -Pattern $Pattern)
    if (-not $files) {
        Write-Log -Level WARNING -Message "Aucun fichier '$Pattern' dans '$SPFolderPath'."
        return [PSCustomObject]@{ Success = 0; Failure = 0; Errors = @() }
    }

    Write-Log -Level INFO -Message "$($files.Count) fichier(s) à télécharger (pattern: $Pattern)"
    $ok = 0; $ko = 0; $errors = [System.Collections.Generic.List[object]]::new()

    foreach ($f in $files) {
        try {
            Invoke-SPDownloadFile -SPFolderPath $SPFolderPath -FileName $f.Name `
                -DestinationPath $DestinationPath -PostProcess $PostProcess -SPArchivePath $SPArchivePath
            $ok++
        } catch {
            $ko++
            $errors.Add([PSCustomObject]@{ File = $f.Name; Error = $_.Exception.Message })
            Write-Log -Level ERROR -Message "Échec download '$($f.Name)' : $_" -Exception $_.Exception
        }
    }

    Write-Log -Level INFO -Message "Download terminé : $ok succès, $ko échec(s)"
    return [PSCustomObject]@{ Success = $ok; Failure = $ko; Errors = $errors.ToArray() }
}
#endregion

# ============================================================
#region Listing, suppression, test d'existence
# ============================================================
function Get-SPFileList {
    <#
    .SYNOPSIS Liste les fichiers d'un dossier SharePoint (pagination automatique).

    .PARAMETER SPFolderPath  Chemin du dossier
    .PARAMETER Pattern       Filtre wildcard optionnel (défaut : *)
    .OUTPUTS   PSCustomObject[] avec Name, SizeKB, LastModified, Id
    #>
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)][string]$SPFolderPath,
        [string]$Pattern = '*'
    )

    $driveId  = Get-SPDriveId
    $apiPath  = ConvertTo-SPApiPath -Path $SPFolderPath
    $endpoint = "drives/$driveId/$apiPath/children" + '?$select=id,name,size,lastModifiedDateTime,file,folder&$top=999'
    $all      = [System.Collections.Generic.List[object]]::new()

    Write-Log -Level INFO -Message "Listing : '$SPFolderPath' (pattern: $Pattern)"

    do {
        $resp = Invoke-GraphRequest -Endpoint $endpoint
        foreach ($item in $resp.value) {
            if ($item.file -and $item.name -like $Pattern) {
                $all.Add([PSCustomObject]@{
                    Name         = $item.name
                    SizeKB       = [Math]::Round($item.size / 1KB, 2)
                    LastModified = [DateTime]$item.lastModifiedDateTime
                    Id           = $item.id
                })
            }
        }
        $endpoint = $resp.'@odata.nextLink'
    } while ($endpoint)

    Write-Log -Level INFO -Message "$($all.Count) fichier(s) trouvé(s)"
    return $all.ToArray()
}

function Remove-SPFile {
    <#
    .SYNOPSIS Supprime un fichier dans SharePoint.

    .PARAMETER SPFolderPath  Dossier SharePoint
    .PARAMETER FileName      Nom du fichier à supprimer
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][string]$SPFolderPath,
        [Parameter(Mandatory)][string]$FileName
    )

    $spPath  = "$($SPFolderPath.TrimEnd('/'))/$FileName"
    $apiPath = ConvertTo-SPApiPath -Path $spPath
    $driveId = Get-SPDriveId

    if ($PSCmdlet.ShouldProcess($spPath, 'Supprimer fichier SharePoint')) {
        Write-Log -Level INFO -Message "Suppression SharePoint : '$spPath'"
        Invoke-GraphRequest -Method DELETE -Endpoint "drives/$driveId/$apiPath"
        Write-Log -Level INFO -Message "Fichier supprimé : '$spPath'"
    }
}

function Remove-SPFilesByPattern {
    <#
    .SYNOPSIS Supprime les fichiers d'un dossier SharePoint correspondant à un pattern.

    .PARAMETER SPFolderPath  Dossier SharePoint
    .PARAMETER Pattern       Filtre wildcard (ex: "*.tmp")
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$SPFolderPath,
        [Parameter(Mandatory)][string]$Pattern
    )

    $files = @(Get-SPFileList -SPFolderPath $SPFolderPath -Pattern $Pattern)
    if (-not $files) {
        Write-Log -Level WARNING -Message "Aucun fichier '$Pattern' dans '$SPFolderPath'."
        return [PSCustomObject]@{ Success = 0; Failure = 0 }
    }

    if ($Pattern -eq '*') {
        Write-Log -Level WARNING -Message "Attention : Pattern='*' → TOUS les fichiers du dossier seront supprimés !"
    }

    Write-Log -Level INFO -Message "$($files.Count) fichier(s) à supprimer (pattern: $Pattern)"
    $ok = 0; $ko = 0

    foreach ($f in $files) {
        try {
            Remove-SPFile -SPFolderPath $SPFolderPath -FileName $f.Name
            $ok++
        } catch {
            $ko++
            Write-Log -Level ERROR -Message "Échec suppression '$($f.Name)' : $_" -Exception $_.Exception
        }
    }

    Write-Log -Level INFO -Message "Suppression terminée : $ok succès, $ko échec(s)"
    return [PSCustomObject]@{ Success = $ok; Failure = $ko }
}

function Test-SPFileExists {
    <#
    .SYNOPSIS Teste l'existence d'un fichier dans un dossier SharePoint.
    Retourne $true si le fichier existe, $false sinon.

    .PARAMETER SPFolderPath  Dossier SharePoint
    .PARAMETER FileName      Nom du fichier
    #>
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$SPFolderPath,
        [Parameter(Mandatory)][string]$FileName
    )

    $spPath  = "$($SPFolderPath.TrimEnd('/'))/$FileName"
    $apiPath = ConvertTo-SPApiPath -Path $spPath
    $driveId = Get-SPDriveId

    try {
        $item   = Invoke-GraphRequest -Endpoint "drives/$driveId/$apiPath"
        $exists = $null -ne $item -and $null -ne $item.file
        Write-Log -Level INFO -Message "Test existence '$spPath' : $exists"
        return $exists
    } catch {
        $code = 0
        try { if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode } } catch {}
        if ($code -eq 404) {
            Write-Log -Level INFO -Message "Test existence '$spPath' : False (404)"
            return $false
        }
        throw
    }
}
#endregion

# ============================================================
#region Point d'entrée — exécution directe uniquement
# ============================================================
if ($MyInvocation.InvocationName -ne '.' -and -not [string]::IsNullOrWhiteSpace($Action)) {

    Import-SPConfig -ConfigFile $ConfigFile
    Write-Log -Level INFO -Message "========== SharePoint Manager | Action: $Action =========="

    switch ($Action) {

        'UploadFile' {
            if (-not $LocalPath) { throw '-LocalPath requis pour UploadFile' }
            if (-not $SPFolder)  { throw '-SPFolder requis pour UploadFile' }
            Invoke-SPUploadFile -LocalFilePath $LocalPath -SPFolderPath $SPFolder `
                -PostProcess $PostProcess -ArchivePath $ArchivePath
        }

        'UploadDirectory' {
            if (-not $LocalDirectory) { throw '-LocalDirectory requis pour UploadDirectory' }
            if (-not $SPFolder)       { throw '-SPFolder requis pour UploadDirectory' }
            Invoke-SPUploadDirectory -LocalDirectory $LocalDirectory -Pattern $Pattern `
                -SPFolderPath $SPFolder -PostProcess $PostProcess -ArchivePath $ArchivePath
        }

        'DownloadFile' {
            if (-not $FileName)        { throw '-FileName requis pour DownloadFile' }
            if (-not $SPFolder)        { throw '-SPFolder requis pour DownloadFile' }
            if (-not $DestinationPath) { throw '-DestinationPath requis pour DownloadFile' }
            Invoke-SPDownloadFile -SPFolderPath $SPFolder -FileName $FileName `
                -DestinationPath $DestinationPath -PostProcess $PostProcess -SPArchivePath $ArchivePath
        }

        'DownloadDirectory' {
            if (-not $SPFolder)        { throw '-SPFolder requis pour DownloadDirectory' }
            if (-not $DestinationPath) { throw '-DestinationPath requis pour DownloadDirectory' }
            Invoke-SPDownloadDirectory -SPFolderPath $SPFolder -Pattern $Pattern `
                -DestinationPath $DestinationPath -PostProcess $PostProcess -SPArchivePath $ArchivePath
        }

        'ListFiles' {
            if (-not $SPFolder) { throw '-SPFolder requis pour ListFiles' }
            $result = Get-SPFileList -SPFolderPath $SPFolder -Pattern $Pattern
            $result | Format-Table -AutoSize Name, SizeKB, LastModified
        }

        'DeleteFile' {
            if (-not $FileName) { throw '-FileName requis pour DeleteFile' }
            if (-not $SPFolder) { throw '-SPFolder requis pour DeleteFile' }
            Remove-SPFile -SPFolderPath $SPFolder -FileName $FileName
        }

        'DeleteByPattern' {
            if (-not $SPFolder)  { throw '-SPFolder requis pour DeleteByPattern' }
            if (-not $Pattern -or $Pattern -eq '*') {
                Write-Log -Level WARNING -Message "Pattern='*' : confirmation requise (-Confirm ou -WhatIf conseillé)"
            }
            Remove-SPFilesByPattern -SPFolderPath $SPFolder -Pattern $Pattern
        }

        'TestFileExists' {
            if (-not $FileName) { throw '-FileName requis pour TestFileExists' }
            if (-not $SPFolder) { throw '-SPFolder requis pour TestFileExists' }
            $exists = Test-SPFileExists -SPFolderPath $SPFolder -FileName $FileName
            Write-Host "Résultat : $exists"
            return $exists
        }
    }

    Write-Log -Level INFO -Message "========== Terminé =========="
}
#endregion
