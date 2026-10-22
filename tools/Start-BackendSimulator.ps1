<#
.SYNOPSIS
    Simulateur local de l'API d'intégration Power BI (développement et recette uniquement).

.DESCRIPTION
    Reproduit le contrat décrit dans docs/INTEGRATION_API.md, comme le fera l'Azure Function de production :

      GET  /api/reports  catalogue des rapports autorisés (rapports de l'espace de travail de test)
      POST /api/embed    { reportId } -> { reportId, embedUrl, embedToken, expiration }

    - un vrai Embed Token en lecture seule est généré auprès de Power BI à CHAQUE appel (chargement et
      renouvellement), avec votre compte au lieu d'un principal de service ;
    - l'en-tête Authorization: Bearer est exigé (401 sinon) ; le jeton n'est pas validé, faute de jeton
      Entra ID en local ;
    - le reportId reçu est contrôlé côté serveur : seuls les rapports de l'espace de travail sont autorisés (404 sinon) ;
    - CORS pour les origines https://*.sharepoint.com et https://localhost, requête préliminaire OPTIONS sans authentification ;
    - erreurs au format { error: { code, message, correlationId } } ;
    - journal des appels dans la console, sans aucun jeton.

    Au premier lancement, l'espace de travail « Visionneuse Power BI - Tests » est créé et le rapport
    d'exemple Microsoft dev/samples/AdventureWorks Sales.pbix (données fictives) y est publié.

    Prérequis : compte Power BI Pro, PPU ou essai gratuit (création d'un espace de travail).
    Arrêt : Ctrl+C.

.PARAMETER Port
    Port d'écoute (7071, port par défaut des Azure Functions en local).

.PARAMETER TokenLifetimeMinutes
    Durée de vie des Embed Tokens générés. 10 permet d'observer un renouvellement au bout d'environ 5 minutes.

.PARAMETER SimulateErrors
    Ajoute au catalogue des rapports de test renvoyant 403, 404, 503 ou retirant l'accès au premier renouvellement.

.PARAMETER PbixPath
    Publie (ou remplace, sous le même identifiant) ce fichier .pbix dans l'espace de travail de test avant de démarrer.

.EXAMPLE
    ./tools/Start-BackendSimulator.ps1
    ./tools/Start-BackendSimulator.ps1 -TokenLifetimeMinutes 10 -SimulateErrors
    ./tools/Start-BackendSimulator.ps1 -PbixPath "C:\Rapports\MonRapport.pbix"
#>
[CmdletBinding()]
param(
    [int] $Port = 7071,
    [ValidateRange(5, 600)] [int] $TokenLifetimeMinutes = 60,
    [switch] $SimulateErrors,
    [string] $PbixPath
)

$ErrorActionPreference = 'Stop'
# Masque les traces MSAL du module Power BI, très verbeuses.
$WarningPreference = 'SilentlyContinue'

$WorkspaceName = 'Visionneuse Power BI - Tests'
$SamplePbixPath = Join-Path $PSScriptRoot '..\dev\samples\AdventureWorks Sales.pbix'
$SamplePbixUrl = 'https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/AdventureWorks%20Sales%20Sample/AdventureWorks%20Sales.pbix'
$AllowedOrigin = '^https://([a-z0-9-]+\.sharepoint\.com|localhost(:\d+)?)$'

$TestReports = [ordered]@{
    'test-acces-refuse'  = '[Test] Accès refusé (403)'
    'test-rapport-supprime' = '[Test] Rapport supprimé (404)'
    'test-indisponible'  = '[Test] Service indisponible (503)'
    'test-perte-droits'  = '[Test] Droits retirés au renouvellement (403)'
}
$CallCounts = @{}

# --- Connexion et préparation de l'espace de travail -------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
    Write-Host 'Installation du module MicrosoftPowerBIMgmt (portée utilisateur)...'
    Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
}
Import-Module MicrosoftPowerBIMgmt.Profile

Write-Host 'Connexion à Power BI (fenêtre de connexion Microsoft)...'
Connect-PowerBIServiceAccount | Out-Null

function Get-WorkspaceReports {
    (Invoke-PowerBIRestMethod -Method Get -Url "groups/$($script:Workspace.id)/reports" | ConvertFrom-Json).value
}

# Import par l'API REST (multipart) plutôt que New-PowerBIReport, qui masque le motif des refus.
function Publish-Pbix([string] $Path) {
    $fullPath = (Resolve-Path $Path).Path
    $name = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
    Write-Host "Publication de « $name » dans « $WorkspaceName » (peut prendre une minute)..."

    $boundary = [guid]::NewGuid().ToString()
    $encoding = [System.Text.Encoding]::UTF8
    $head = $encoding.GetBytes("--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$name.pbix`"`r`nContent-Type: application/octet-stream`r`n`r`n")
    $tail = $encoding.GetBytes("`r`n--$boundary--`r`n")
    $file = [System.IO.File]::ReadAllBytes($fullPath)
    $body = New-Object byte[] ($head.Length + $file.Length + $tail.Length)
    [Array]::Copy($head, 0, $body, 0, $head.Length)
    [Array]::Copy($file, 0, $body, $head.Length, $file.Length)
    [Array]::Copy($tail, 0, $body, $head.Length + $file.Length, $tail.Length)

    $headers = @{ Authorization = (Get-PowerBIAccessToken -AsString) }
    $baseUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($script:Workspace.id)/imports"
    try {
        $import = Invoke-RestMethod -Method Post -Uri "$($baseUrl)?datasetDisplayName=$([uri]::EscapeDataString($name)).pbix&nameConflict=CreateOrOverwrite" `
            -Headers $headers -ContentType "multipart/form-data; boundary=$boundary" -Body $body
        do {
            Start-Sleep -Seconds 3
            $import = Invoke-RestMethod -Method Get -Uri "$baseUrl/$($import.id)" -Headers $headers
        } while ($import.importState -eq 'Publishing')
    } catch {
        $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Host "Publication refusée par Power BI : $detail" -ForegroundColor Red
        if ($detail -match 'PowerBINotLicensedException') {
            Write-Host @"
Votre compte n'a pas de licence Power BI Pro ou PPU (licence gratuite) : publication et Embed Tokens impossibles.
Activez l'essai gratuit (app.powerbi.com > photo de profil ou Paramètres > Gérer la licence personnelle > Démarrer l'essai),
ou demandez une licence Pro à votre administrateur, puis relancez le simulateur.
"@ -ForegroundColor Yellow
            throw 'Licence Power BI Pro ou PPU requise.'
        }
        Write-Host @"
Causes fréquentes :
  - licence Power BI gratuite : une licence Pro, PPU ou un essai est nécessaire (app.powerbi.com > Paramètres > Gérer la licence personnelle) ;
  - l'administrateur Power BI du tenant bloque l'import de fichiers ou la création d'espaces de travail.
Contournement : dans app.powerbi.com, espace « $WorkspaceName » > Importer > Rapport > À partir de cet ordinateur,
choisir $fullPath, puis relancer le simulateur (il réutilise un rapport déjà publié).
"@ -ForegroundColor Yellow
        throw
    }
    if ($import.importState -ne 'Succeeded') {
        throw "Import Power BI terminé avec l'état « $($import.importState) »."
    }
}

$script:Workspace = (Invoke-PowerBIRestMethod -Method Get -Url 'groups' | ConvertFrom-Json).value |
    Where-Object { $_.name -eq $WorkspaceName } | Select-Object -First 1
if (-not $script:Workspace) {
    Write-Host "Création de l'espace de travail « $WorkspaceName »..."
    $script:Workspace = Invoke-PowerBIRestMethod -Method Post -Url 'groups?workspaceV2=True' -Body (@{ name = $WorkspaceName } | ConvertTo-Json) | ConvertFrom-Json
}

if ($PbixPath) {
    Publish-Pbix $PbixPath
}
if (-not (Get-WorkspaceReports | Where-Object { $_.name -eq 'AdventureWorks Sales' })) {
    if (-not (Test-Path $SamplePbixPath)) {
        Write-Host "Téléchargement du rapport d'exemple Microsoft (8 Mo)..."
        New-Item -ItemType Directory -Force -Path (Split-Path $SamplePbixPath) | Out-Null
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $SamplePbixUrl -OutFile $SamplePbixPath -UseBasicParsing
    }
    Publish-Pbix $SamplePbixPath
}

# --- Réponses HTTP ------------------------------------------------------------------------------------------

function Write-Log([string] $Method, [string] $Path, [int] $Status, [string] $Detail) {
    $color = if ($Status -ge 400) { 'Yellow' } else { 'Gray' }
    Write-Host ('{0:HH:mm:ss}  {1,-7} {2,-13} {3}  {4}' -f (Get-Date), $Method, $Path, $Status, $Detail) -ForegroundColor $color
}

function Send-Json($Context, [int] $Status, $Body, [string] $Detail) {
    $response = $Context.Response
    $response.StatusCode = $Status
    $response.ContentType = 'application/json; charset=utf-8'
    $response.AddHeader('Cache-Control', 'no-store')
    $bytes = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Body -Depth 5 -Compress))
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.Close()
    Write-Log $Context.Request.HttpMethod $Context.Request.Url.AbsolutePath $Status $Detail
}

function Send-Error($Context, [int] $Status, [string] $Code, [string] $Message) {
    $correlationId = [guid]::NewGuid().ToString()
    Send-Json $Context $Status @{ error = @{ code = $Code; message = $Message; correlationId = $correlationId } } "$Code (correlationId $correlationId)"
}

function Get-Catalog {
    $catalog = @(Get-WorkspaceReports | ForEach-Object { [ordered]@{ id = $_.id; name = $_.name } })
    if ($SimulateErrors) {
        $catalog += @($TestReports.Keys | ForEach-Object { [ordered]@{ id = $_; name = $TestReports[$_] } })
    }
    , $catalog
}

# Équivalent de ce que fera l'Azure Function : embedUrl du rapport + GenerateToken en lecture seule.
function New-EmbedResponse($Report) {
    $body = @{ accessLevel = 'View'; lifetimeInMinutes = $TokenLifetimeMinutes } | ConvertTo-Json
    $raw = Invoke-PowerBIRestMethod -Method Post -Url "groups/$($script:Workspace.id)/reports/$($Report.id)/GenerateToken" -Body $body
    # Extraction brute : ConvertFrom-Json (PowerShell 7) convertirait la date en heure locale sans fuseau.
    [ordered]@{
        reportId   = $Report.id
        embedUrl   = $Report.embedUrl
        embedToken = [regex]::Match($raw, '"token"\s*:\s*"([^"]+)"').Groups[1].Value
        expiration = [regex]::Match($raw, '"expiration"\s*:\s*"([^"]+)"').Groups[1].Value
    }
}

function Invoke-Embed($Context) {
    $reader = New-Object System.IO.StreamReader($Context.Request.InputStream, [System.Text.Encoding]::UTF8)
    $raw = $reader.ReadToEnd()
    try { $payload = $raw | ConvertFrom-Json } catch { $payload = $null }
    $reportId = if ($payload) { $payload.reportId }
    if (-not ($reportId -is [string]) -or -not $reportId.Trim() -or $reportId.Length -gt 200) {
        return Send-Error $Context 400 'INVALID_REQUEST' 'reportId manquant ou invalide'
    }

    $reports = @(Get-WorkspaceReports)
    $sample = $reports | Where-Object { $_.name -eq 'AdventureWorks Sales' } | Select-Object -First 1
    if ($SimulateErrors -and $TestReports.Contains($reportId)) {
        $CallCounts[$reportId] = 1 + [int]$CallCounts[$reportId]
        switch ($reportId) {
            'test-acces-refuse' { return Send-Error $Context 403 'REPORT_FORBIDDEN' "Utilisateur non autorisé sur ce rapport" }
            'test-rapport-supprime' { return Send-Error $Context 404 'REPORT_NOT_FOUND' 'Rapport supprimé dans Power BI' }
            'test-indisponible' { return Send-Error $Context 503 'SERVICE_UNAVAILABLE' 'Indisponibilité simulée' }
            'test-perte-droits' {
                if ($CallCounts[$reportId] -gt 1) {
                    return Send-Error $Context 403 'USER_FORBIDDEN' "Utilisateur retiré du groupe autorisé"
                }
                return Send-Json $Context 200 (New-EmbedResponse $sample) "jeton généré (AdventureWorks Sales, accès retiré au prochain appel)"
            }
        }
    }

    # Contrôle côté serveur : le reportId du navigateur n'est accepté que s'il figure dans la liste autorisée.
    $report = $reports | Where-Object { $_.id -eq $reportId } | Select-Object -First 1
    if (-not $report) {
        return Send-Error $Context 404 'REPORT_NOT_FOUND' 'Rapport inconnu ou non autorisé'
    }
    $result = New-EmbedResponse $report
    Send-Json $Context 200 $result "jeton généré ($($report.name), expire $($result.expiration))"
}

function Invoke-Request($Context) {
    $request = $Context.Request
    $origin = $request.Headers['Origin']
    if ($origin -and $origin -match $AllowedOrigin) {
        $Context.Response.AddHeader('Access-Control-Allow-Origin', $origin)
        $Context.Response.AddHeader('Vary', 'Origin')
        $Context.Response.AddHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        $Context.Response.AddHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type, Accept')
        $Context.Response.AddHeader('Access-Control-Allow-Private-Network', 'true')
    }

    # Requête préliminaire CORS : jamais d'authentification.
    if ($request.HttpMethod -eq 'OPTIONS') {
        $Context.Response.StatusCode = 204
        $Context.Response.Close()
        Write-Log 'OPTIONS' $request.Url.AbsolutePath 204 "préliminaire CORS ($origin)"
        return
    }
    if ($request.Headers['Authorization'] -notmatch '^Bearer\s+\S+') {
        return Send-Error $Context 401 'UNAUTHENTICATED' 'En-tête Authorization: Bearer absent'
    }

    switch ("$($request.HttpMethod) $($request.Url.AbsolutePath.TrimEnd('/'))") {
        'GET /api/reports' { $catalog = Get-Catalog; Send-Json $Context 200 $catalog "$($catalog.Count) rapport(s)" }
        'POST /api/embed' { Invoke-Embed $Context }
        default { Send-Error $Context 404 'ROUTE_NOT_FOUND' 'Route inconnue' }
    }
}

# --- Boucle d'écoute ----------------------------------------------------------------------------------------

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()

Write-Host ''
Write-Host "Simulateur de l'API d'intégration à l'écoute sur http://localhost:$Port" -ForegroundColor Green
Write-Host "Espace de travail : $WorkspaceName ; durée des jetons : $TokenLifetimeMinutes min"
foreach ($item in (Get-Catalog)) { Write-Host "  - $($item.name)  [$($item.id)]" }
Write-Host 'Workbench : propriétés de la WebPart > Mode simulation activé > choisir un rapport. Arrêt : Ctrl+C.'
Write-Host ''

try {
    while ($listener.IsListening) {
        $pending = $listener.GetContextAsync()
        # Attente par tranches pour que Ctrl+C reste opérant.
        while (-not $pending.Wait(500)) { }
        $context = $pending.Result
        try {
            Invoke-Request $context
        } catch {
            Write-Host "Erreur lors du traitement : $($_.Exception.Message)" -ForegroundColor Red
            try { Send-Error $context 500 'POWERBI_ERROR' 'Erreur lors de l''appel à Power BI' } catch { }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
    Write-Host 'Simulateur arrêté.'
}
