#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    SharePoint File Management via module Microsoft.Graph

.DESCRIPTION
    Gestion de fichiers SharePoint via le module Microsoft.Graph.
    Authentification par certificat (Connect-MgGraph).
    Simplifié par rapport à la version API pure : pas de gestion de JWT ni de token.

.PARAMETER Action
    UploadFile | UploadDirectory | DownloadFile | DownloadDirectory |
    ListFiles | DeleteFile | DeleteByPattern | TestFileExists

.PARAMETER LocalPath        Chemin du fichier local (upload/download unitaire)
.PARAMETER LocalDirectory   Répertoire local (upload/download en lot)
.PARAMETER SPFolder         Chemin du dossier SharePoint (ex: "Documents/Rapports")
.PARAMETER FileName         Nom du fichier dans SharePoint
.PARAMETER Pattern          Filtre wildcard (ex: "*.csv")  [défaut: *]
.PARAMETER DestinationPath  Répertoire de destination local (download)
.PARAMETER PostProcess      None | Delete | Archive  — traitement du source après succès
.PARAMETER ArchivePath      Répertoire local d'archivage (si PostProcess = Archive)

.EXAMPLE
    .\SharePoint-Manager-MgGraph.ps1 -Action UploadFile -LocalPath "C:\data\rapport.xlsx" -SPFolder "Documents/Rapports"

.EXAMPLE
    .\SharePoint-Manager-MgGraph.ps1 -Action UploadDirectory -LocalDirectory "C:\exports" -Pattern "*.csv" `
        -SPFolder "Data/Imports" -PostProcess Archive -ArchivePath "C:\done"

.EXAMPLE
    .\SharePoint-Manager-MgGraph.ps1 -Action ListFiles -SPFolder "Documents/Rapports" -Pattern "2024_*"

.NOTES
    Prérequis :
      Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
      - Enregistrement d'application Azure AD avec permission Sites.ReadWrite.All (application)
      - Certificat chargé dans l'enregistrement d'application
      - Certificat (avec clé privée) installé dans CurrentUser\My ou LocalMachine\My
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
    [string]$ArchivePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
#region Configuration  —  ADAPTER AVANT UTILISATION
# ============================================================
$Script:Config = [ordered]@{
    # Azure AD / Enregistrement d'application
    TenantId           = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    ClientId           = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    CertThumbprint     = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

    # SharePoint
    SharePointHost     = 'votre-tenant.sharepoint.com'
    SitePath           = '/sites/votre-site'
    DriveName          = 'Documents'    # vide = drive par défaut

    # Comportement
    LargeFileThreshold = 4MB
    UploadChunkSize    = 10485760       # 10 Mo (multiple de 327 680)

    # Journalisation
    LogFile            = (Join-Path $PSScriptRoot 'SharePoint-Manager.log')
    LogLevel           = 'INFO'         # DEBUG | INFO | WARNING | ERROR
    LogToConsole       = $true

    # URL de base Graph (ne pas modifier)
    GraphBaseUrl       = 'https://graph.microsoft.com/v1.0'
}

# Caches internes
$Script:CachedSiteId  = $null
$Script:CachedDriveId = $null
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

    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
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
#region Authentification  —  Connect-MgGraph gère tout
# ============================================================
function Connect-SPGraph {
    <#
    .SYNOPSIS Ouvre une session Graph si aucune session valide n'existe déjà.
    #>
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if ($ctx -and $ctx.ClientId -eq $Script:Config.ClientId -and $ctx.TenantId -eq $Script:Config.TenantId) {
        Write-Log -Level DEBUG -Message 'Session Graph déjà active.'
        return
    }

    Write-Log -Level INFO -Message 'Connexion à Microsoft Graph (certificat)...'
    Connect-MgGraph -TenantId      $Script:Config.TenantId `
                    -ClientId      $Script:Config.ClientId `
                    -CertificateThumbprint $Script:Config.CertThumbprint `
                    -NoWelcome -ErrorAction Stop
    Write-Log -Level INFO -Message 'Connecté à Microsoft Graph.'
}
#endregion

# ============================================================
#region Helpers SharePoint
# ============================================================
function Invoke-SP {
    <#
    .SYNOPSIS Raccourci pour Invoke-MgGraphRequest avec l'URL Graph de base.
    #>
    param(
        [Parameter(Mandatory)][string]$Endpoint,
        [ValidateSet('GET','POST','PUT','PATCH','DELETE')]
        [string]$Method = 'GET',
        [object]$Body         = $null,
        [string]$ContentType  = 'application/json',
        [string]$OutputFilePath
    )

    $uri    = if ($Endpoint -match '^https?://') { $Endpoint } `
              else { "$($Script:Config.GraphBaseUrl)/$Endpoint" }
    $params = @{ Uri = $uri; Method = $Method; ErrorAction = 'Stop' }

    if ($Body) {
        $params.Body        = $Body
        $params.ContentType = $ContentType
    }
    if ($OutputFilePath) { $params.OutputFilePath = $OutputFilePath }

    return Invoke-MgGraphRequest @params
}

function Get-SPSiteId {
    [OutputType([string])] param()
    if ($Script:CachedSiteId) { return $Script:CachedSiteId }

    $h = $Script:Config.SharePointHost
    $p = $Script:Config.SitePath.TrimStart('/')
    $Script:CachedSiteId = (Invoke-SP -Endpoint "sites/${h}:/${p}").id
    Write-Log -Level DEBUG -Message "Site ID : $($Script:CachedSiteId)"
    return $Script:CachedSiteId
}

function Get-SPDriveId {
    [OutputType([string])] param()
    if ($Script:CachedDriveId) { return $Script:CachedDriveId }

    $siteId    = Get-SPSiteId
    $driveName = $Script:Config.DriveName

    if ([string]::IsNullOrWhiteSpace($driveName)) {
        $Script:CachedDriveId = (Invoke-SP -Endpoint "sites/$siteId/drive").id
    } else {
        $drives = (Invoke-SP -Endpoint "sites/$siteId/drives").value
        $drive  = $drives | Where-Object { $_.name -eq $driveName }
        if (-not $drive) { throw "Bibliothèque '$driveName' introuvable. Disponibles : $(($drives.name) -join ', ')" }
        $Script:CachedDriveId = $drive.id
    }

    Write-Log -Level DEBUG -Message "Drive ID : $($Script:CachedDriveId)"
    return $Script:CachedDriveId
}

function ConvertTo-SPApiPath {
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
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('None','Delete','Archive')][string]$Action = 'None',
        [string]$ArchivePath
    )
    switch ($Action) {
        'None'    { return }
        'Delete'  {
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
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$SPFullPath,
        [Parameter(Mandatory)][string]$DriveId
    )

    $fi        = Get-Item $FilePath
    $apiPath   = ConvertTo-SPApiPath -Path $SPFullPath
    $chunkSize = $Script:Config.UploadChunkSize

    Write-Log -Level DEBUG -Message "Session upload '$($fi.Name)' ($([Math]::Round($fi.Length/1MB,2)) Mo)"

    $session   = Invoke-SP -Method POST `
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
            $chunk = if ($read -eq $chunkSize) { $buffer } else { $buffer[0..($read - 1)] }
            $end   = $offset + $read - 1
            $pct   = [Math]::Round(($offset / $total) * 100, 1)
            Write-Progress -Activity "Upload $($fi.Name)" -Status "$pct%" -PercentComplete $pct

            # L'URL de session est pré-autorisée : pas besoin de passer par Invoke-MgGraphRequest
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
            Invoke-SP -Method PUT -Endpoint "drives/$driveId/$apiPath/content" `
                -Body ([IO.File]::ReadAllBytes($fi.FullName)) -ContentType 'application/octet-stream'
        }

        Write-Log -Level INFO -Message "Upload réussi : $($result.name)"
        Invoke-LocalPostProcess -FilePath $fi.FullName -Action $PostProcess -ArchivePath $ArchivePath
        return $result
    }
}

function Invoke-SPUploadDirectory {
    <#
    .SYNOPSIS Upload des fichiers d'un répertoire local correspondant à un pattern.
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
    .SYNOPSIS Télécharge un fichier depuis SharePoint.
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

        # OutputFilePath déclenche un téléchargement binaire dans Invoke-MgGraphRequest
        Invoke-SP -Endpoint "drives/$driveId/$apiPath/content" -OutputFilePath $outFile
        Write-Log -Level INFO -Message "Download réussi : '$outFile'"

        switch ($PostProcess) {
            'Delete' {
                Write-Log -Level INFO -Message "Post-process : suppression SharePoint '$spPath'"
                Remove-SPFile -SPFolderPath $SPFolderPath -FileName $FileName
            }
            'Archive' {
                if (-not $SPArchivePath) { throw "-SPArchivePath requis avec PostProcess=Archive" }
                $item       = Invoke-SP -Endpoint "drives/$driveId/$apiPath"
                $destFolder = Invoke-SP -Endpoint "drives/$driveId/$(ConvertTo-SPApiPath $SPArchivePath)"
                Invoke-SP -Method PATCH -Endpoint "drives/$driveId/items/$($item.id)" `
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
        $resp = Invoke-SP -Endpoint $endpoint
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
        Invoke-SP -Method DELETE -Endpoint "drives/$driveId/$apiPath"
        Write-Log -Level INFO -Message "Fichier supprimé."
    }
}

function Remove-SPFilesByPattern {
    <#
    .SYNOPSIS Supprime les fichiers d'un dossier SharePoint correspondant à un pattern.
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
        Write-Log -Level WARNING -Message "Pattern='*' → TOUS les fichiers du dossier seront supprimés !"
    }

    Write-Log -Level INFO -Message "$($files.Count) fichier(s) à supprimer"
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
        $item   = Invoke-SP -Endpoint "drives/$driveId/$apiPath"
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

    Connect-SPGraph

    Write-Log -Level INFO -Message "========== SharePoint Manager (MgGraph) | Action: $Action =========="

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
            Get-SPFileList -SPFolderPath $SPFolder -Pattern $Pattern |
                Format-Table -AutoSize Name, SizeKB, LastModified
        }

        'DeleteFile' {
            if (-not $FileName) { throw '-FileName requis pour DeleteFile' }
            if (-not $SPFolder) { throw '-SPFolder requis pour DeleteFile' }
            Remove-SPFile -SPFolderPath $SPFolder -FileName $FileName
        }

        'DeleteByPattern' {
            if (-not $SPFolder) { throw '-SPFolder requis pour DeleteByPattern' }
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
