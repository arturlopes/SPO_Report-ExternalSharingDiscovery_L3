<#
.SYNOPSIS
  SPO External Access v1 (Web + Libraries + Folders) + Progress — QuickOps
.DESCRIPTION
  Lê targets (2 sites/subsites) e identifica usuários externos com acesso em:
   - Site/Subsite (Web)
   - Bibliotecas (Document Libraries)
   - Pastas (somente FSObjType=1) COM herança quebrada (HasUniqueRoleAssignments=True)
  Externo = Guest (#ext#/urn:spo:guest) OU domínio de e-mail fora dos domínios corporativos.

  Sem arquivos (file-level) por solicitação.
.OUTPUT
  01_Summary_<runId>.csv
  02_TargetResolution_<runId>.csv
  03_WebExternalAccess_<runId>.csv
  04_LibraryExternalAccess_<runId>.csv
  05_FolderExternalAccess_<runId>.csv
  06_Skipped_<runId>.csv
  + XLSX opcional + Transcript
#>

# ============================= CONFIG =============================
$Tenant      = "5ef881c6-1fa8-4d9c-ad81-634a9d5f9d07"
$ClientId    = "fa70bcfc-402a-4156-8634-c8b8cd772bfb"
$Thumbprint  = "EEED6DA026421085D8DB922A7973EE913CE26645"

$TargetsFile = "C:\Temp\Input\sites.txt"
$OutputRoot  = "C:\Temp\Output\SPO_ExternalAccess_v1"

$IncludeSystemLists = $false

# Domínios corporativos (para detectar externo por domínio)
$CorporateDomains = @("")

# Expandir SharePoint Groups (best-effort; não usa Graph)
$ExpandSharePointGroups = $true

# Pastas: apenas com herança quebrada (recomendado para performance)
$FoldersOnlyUniqueAcl = $false

# PageSize para listagem de itens (pastas)
$PageSize = 2000

# ============================= VALIDATION / PREREQ =============================
if (-not (Test-Path $TargetsFile)) { throw "TargetsFile não encontrado: $TargetsFile" }

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Import-Module PnP.PowerShell -ErrorAction Stop

$HasImportExcel = $false
try { Import-Module ImportExcel -ErrorAction Stop; $HasImportExcel = $true } catch { $HasImportExcel = $false }

# ============================= HELPERS =============================
function Write-QO {
  param([Parameter(Mandatory)][string]$Message,[ValidateSet("INFO","OK","WARN","ERROR")][string]$Level="INFO")
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $color = @{INFO="Cyan";OK="Green";WARN="Yellow";ERROR="Red"}[$Level]
  Write-Host "[$ts] [$($Level.PadRight(5))] $Message" -ForegroundColor $color
}

function Normalize-Url {
  param([Parameter(Mandatory)][string]$Url)

  $raw = $Url.Trim().Trim('"').Trim("'")
  if (-not $raw) { return $null }
  if ($raw -notmatch '^https?://') { $raw = "https://$raw" }

  # Remove query/fragment e páginas AllItems.aspx / _layouts
  try {
    $u = [Uri]$raw
    $clean = "$($u.Scheme)://$($u.Host)$($u.AbsolutePath)"
  } catch {
    $clean = $raw
  }

  $clean = $clean -replace '/Forms/AllItems\.aspx$',''
  $clean = $clean -replace '/_layouts/15/.*$',''

  return $clean.TrimEnd('/')
}

function Export-CsvUtf8Bom {
  param([Parameter(Mandatory)]$Data,[Parameter(Mandatory)][string]$Path)
  $folder = Split-Path $Path -Parent
  if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
  $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8
}

function Connect-QOPnP {
  param([Parameter(Mandatory)][string]$Url)
  Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId -Thumbprint $Thumbprint -ErrorAction Stop
}

function Get-RootSiteUrl {
  param([Parameter(Mandatory)][string]$TargetUrl)
  $uri  = [Uri]$TargetUrl
  $path = $uri.AbsolutePath.TrimEnd('/')
  $m = [regex]::Match($path, '^/(sites|teams)/[^/]+', 'IgnoreCase')
  if ($m.Success) { return "$($uri.Scheme)://$($uri.Host)$($m.Value)" }
  return "$($uri.Scheme)://$($uri.Host)"
}

function Get-TenantHostUrl {
  param([Parameter(Mandatory)][string]$AnyWebUrl)
  $u = [Uri]$AnyWebUrl
  return "$($u.Scheme)://$($u.Host)"
}

function Get-FullUrlFromServerRelative {
  param([Parameter(Mandatory)][string]$TenantHostUrl,[Parameter(Mandatory)][string]$ServerRelativeUrl)
  $base = $TenantHostUrl.TrimEnd('/')
  $rel  = $ServerRelativeUrl.TrimStart('/')
  return "$base/$rel"
}

function Test-IsExternalPrincipal {
  param([string]$LoginName,[string]$Email)

  if ($LoginName -match '#ext#') { return $true }
  if ($LoginName -match 'urn:spo:guest') { return $true }

  if ($Email -and $CorporateDomains.Count -gt 0) {
    $parts = $Email.Split('@')
    if ($parts.Count -eq 2) {
      $dom = $parts[1].ToLower()
      if ($dom -and ($CorporateDomains -notcontains $dom)) { return $true }
    }
  }
  return $false
}

# ============================= DISCOVERY =============================
function Get-SubWebsDeep {
  param([Parameter(Mandatory)][string]$RootUrl)

  Connect-QOPnP -Url $RootUrl
  $rows = New-Object System.Collections.Generic.List[object]

  $rootWeb = Get-PnPWeb -Includes Title, Url, ServerRelativeUrl -ErrorAction Stop
  $rows.Add([pscustomobject]@{Title=$rootWeb.Title; Url=$rootWeb.Url.TrimEnd('/'); ServerRelativeUrl=$rootWeb.ServerRelativeUrl.TrimEnd('/'); IsRoot=$true}) | Out-Null

  $subs = Get-PnPSubWeb -Recurse -Includes Title, Url, ServerRelativeUrl -ErrorAction Stop
  foreach ($s in $subs) {
    $rows.Add([pscustomobject]@{Title=$s.Title; Url=$s.Url.TrimEnd('/'); ServerRelativeUrl=$s.ServerRelativeUrl.TrimEnd('/'); IsRoot=$false}) | Out-Null
  }
  return $rows
}

function Resolve-TargetToBranchWeb {
  param([Parameter(Mandatory)][string]$TargetUrl,[Parameter(Mandatory)]$SubWebs)
  $target = $TargetUrl.TrimEnd('/')
  $best = $null; $bestLen = -1
  foreach ($sw in $SubWebs) {
    if ($target.StartsWith($sw.Url, [System.StringComparison]::OrdinalIgnoreCase)) {
      if ($sw.Url.Length -gt $bestLen) { $best = $sw; $bestLen = $sw.Url.Length }
    }
  }
  if (-not $best) { return [pscustomobject]@{Found=$false; BranchWebUrl=""; BranchWebTitle=""; Note="Nenhum subweb compatível"} }
  return [pscustomobject]@{Found=$true; BranchWebUrl=$best.Url; BranchWebTitle=$best.Title; Note="Resolved by longest-prefix"} 
}

function Get-WebBranchSet {
  param([Parameter(Mandatory)]$AllWebs,[Parameter(Mandatory)][string]$BranchWebUrl)
  return $AllWebs | Where-Object { $_.Url.StartsWith($BranchWebUrl,[System.StringComparison]::OrdinalIgnoreCase) }
}

# ============================= ACL =============================
function Expand-RoleAssignmentsBase {
  param([Parameter(Mandatory)]$SecurableObject)

  $rows = New-Object System.Collections.Generic.List[object]

  Get-PnPProperty -ClientObject $SecurableObject -Property HasUniqueRoleAssignments, RoleAssignments | Out-Null
  $hasUnique = [bool]$SecurableObject.HasUniqueRoleAssignments

  foreach ($ra in $SecurableObject.RoleAssignments) {
    try {
      Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
      $m = $ra.Member
      $roles = ($ra.RoleDefinitionBindings | ForEach-Object { $_.Name }) -join '; '

      $rows.Add([pscustomobject]@{
        PrincipalTitle = $m.Title
        PrincipalLogin = $m.LoginName
        PrincipalEmail = $m.Email
        PrincipalType  = $m.PrincipalType
        Roles          = $roles
      }) | Out-Null
    } catch { continue }
  }

  return [pscustomobject]@{HasUnique=$hasUnique; Rows=$rows}
}

function Expand-RoleAssignmentsWithSPGroupExpansion {
  param(
    [Parameter(Mandatory)]$SecurableObject,
    [Parameter(Mandatory)][string]$ScopeUrl,
    [Parameter(Mandatory)]$SkippedOut
  )

  $base = Expand-RoleAssignmentsBase -SecurableObject $SecurableObject
  $out  = New-Object System.Collections.Generic.List[object]

  foreach ($row in $base.Rows) {
    $ptype = ""
    try { $ptype = $row.PrincipalType.ToString() } catch {}

$isSpGroup = $ExpandSharePointGroups -and (
  $row.PrincipalType -eq "SharePointGroup" -or
  $row.PrincipalLogin -like "*sharepoint*" -or
  $row.PrincipalLogin -like "*Visitors*" -or
  $row.PrincipalLogin -like "*Members*" -or
  $row.PrincipalLogin -like "*Owners*"
)

    if (-not $isSpGroup) {
      $out.Add([pscustomobject]@{
        GrantedThrough = "Direct"
        MemberTitle    = $row.PrincipalTitle
        MemberLogin    = $row.PrincipalLogin
        MemberEmail    = $row.PrincipalEmail
        PrincipalType  = $row.PrincipalType
        Roles          = $row.Roles
        IsExternal     = (Test-IsExternalPrincipal -LoginName $row.PrincipalLogin -Email $row.PrincipalEmail)
      }) | Out-Null
      continue
    }

    # Best-effort: expand SharePoint Group members (sem Graph)
    try {
      Connect-QOPnP -Url $ScopeUrl
      # 🔧 FIX: resolver grupo corretamente por ID
$group = Get-PnPGroup | Where-Object { 
  $_.Title -eq $row.PrincipalTitle -or 
  $_.LoginName -eq $row.PrincipalLogin
}
  Write-QO "Expandindo grupo: $($row.PrincipalTitle)" "INFO"

if ($group) {

  $members = Get-PnPGroupMember -Identity $group.Id -ErrorAction Stop

  foreach ($m in $members) {
    $login=$null; $email=$null; $mtitle=$null; $mtype=$null
    try { $login=$m.LoginName } catch {}
    try { $email=$m.Email } catch {}
    try { $mtitle=$m.Title } catch {}
    try { $mtype=$m.PrincipalType } catch {}

    $out.Add([pscustomobject]@{
      GrantedThrough = "SPGroup: $($row.PrincipalTitle)"
      MemberTitle    = $mtitle
      MemberLogin    = $login
      MemberEmail    = $email
      PrincipalType  = $mtype
      Roles          = $row.Roles
      IsExternal     = (Test-IsExternalPrincipal -LoginName $login -Email $email)
    }) | Out-Null
  }

} else {
  # 🔧 opcional: log de debug
  $SkippedOut.Add([pscustomobject]@{
    Scope="GroupResolution"
    Url=$ScopeUrl
    Object=$row.PrincipalTitle
    Reason="Grupo não encontrado"
  }) | Out-Null
}
   
    } catch {
      $SkippedOut.Add([pscustomobject]@{
        Scope="GroupExpansion"; Url=$ScopeUrl; Object=$row.PrincipalTitle; Reason=$_.Exception.Message
      }) | Out-Null

      # mantém o grupo (não expandido)
      $out.Add([pscustomobject]@{
        GrantedThrough = "Group(Unexpanded)"
        MemberTitle    = $row.PrincipalTitle
        MemberLogin    = $row.PrincipalLogin
        MemberEmail    = $row.PrincipalEmail
        PrincipalType  = $row.PrincipalType
        Roles          = $row.Roles
        IsExternal     = (Test-IsExternalPrincipal -LoginName $row.PrincipalLogin -Email $row.PrincipalEmail)
      }) | Out-Null
    }
  }

  return [pscustomobject]@{HasUnique=$base.HasUnique; Rows=$out}
}

function Get-LibrariesForWeb {
  param([Parameter(Mandatory)][string]$WebUrl)
  Connect-QOPnP -Url $WebUrl

  $libs = Get-PnPList -Includes Title, RootFolder, Hidden, BaseTemplate, HasUniqueRoleAssignments, RoleAssignments |
    Where-Object { $_.BaseTemplate -eq 101 -and ($IncludeSystemLists -or (-not $_.Hidden)) }

  if (-not $IncludeSystemLists) {
    $libs = $libs | Where-Object {
      $_.Title -notin @("Style Library","Site Assets","Site Pages","Form Templates","Site Collection Documents","Site Collection Images") -and
      $_.Title -notmatch '^DO_NOT_DELETE_'
    }
  }
  return $libs
}


## New Function ##

function Get-FolderExternalAclRows {
  param(
    [Parameter(Mandatory)][string]$WebUrl,
    [Parameter(Mandatory)][string]$LibraryTitle,
    [Parameter(Mandatory)][string]$TargetUrl,
    [Parameter(Mandatory)][string]$RootUrl,
    [Parameter(Mandatory)][string]$TenantHostUrl,
    [Parameter(Mandatory)]$SkippedOut,
    [Parameter(Mandatory)]$WebOut,
    [Parameter(Mandatory)][string]$RunId
  )

  $rows = New-Object System.Collections.Generic.List[object]

  try {
    Connect-QOPnP -Url $WebUrl

    # 🔧 FIX 1: garantir campos necessários + evitar lazy loading incompleto
    $items = Get-PnPListItem -List $LibraryTitle -PageSize $PageSize `
      -Fields "ID","FileRef","FileLeafRef","FSObjType","HasUniqueRoleAssignments"

    foreach ($it in $items) {

      # Apenas pastas
      if ([int]$it["FSObjType"] -ne 1) { continue }

      # 🔧 FIX 2: proteção contra FileRef vazio (causava erro de URL)
      $ref = [string]$it["FileRef"]
      if ([string]::IsNullOrWhiteSpace($ref)) { continue }

      # 🔥 HERDAR usuários externos do Web
      $webUsers = $WebOut | Where-Object { $_.WebUrl -eq $WebUrl }

      foreach ($wu in $webUsers) {
        $rows.Add([pscustomobject]@{
          RunId=$RunId
          TargetUrl=$TargetUrl
          RootUrl=$RootUrl
          WebUrl=$WebUrl
          Level="Folder"
          LibraryTitle=$LibraryTitle
          FolderName=[string]$it["FileLeafRef"]
          FolderServerRelative=$ref
          FolderUrl=Get-FullUrlFromServerRelative `
            -TenantHostUrl $TenantHostUrl `
            -ServerRelativeUrl $ref
          InheritanceStatus="Inherited"
          GrantedThrough=$wu.GrantedThrough
          UserLogin=$wu.UserLogin
          UserEmail=$wu.UserEmail
          Roles=$wu.Roles
        }) | Out-Null
      }

      # Apenas Unique ACL (se configurado)
      if ($FoldersOnlyUniqueAcl -and (-not [bool]$it["HasUniqueRoleAssignments"])) { continue }

      $id  = $it["ID"]

      try {
        # 🔧 FIX 3 (CRÍTICO): carregar item COMPLETO com propriedades de segurança
        $fullItem = Get-PnPListItem -List $LibraryTitle -Id $id `
          -Fields "ID","FileRef","FileLeafRef","FSObjType","HasUniqueRoleAssignments" `
          -ErrorAction Stop

        # 🔧 FIX 4 (CRÍTICO): garantir carregamento explícito de RoleAssignments
        Get-PnPProperty -ClientObject $fullItem -Property HasUniqueRoleAssignments, RoleAssignments

        # 🔧 FIX 5: evitar processar sem RoleAssignments carregado
        if (-not $fullItem.RoleAssignments) { continue }

        $acl = Expand-RoleAssignmentsWithSPGroupExpansion `
          -SecurableObject $fullItem `
          -ScopeUrl $WebUrl `
          -SkippedOut $SkippedOut

        foreach ($r in $acl.Rows) {
          if (-not $r.IsExternal) { continue }

          $rows.Add([pscustomobject]@{
            RunId=$runId
            TargetUrl=$TargetUrl
            RootUrl=$RootUrl
            WebUrl=$WebUrl
            Level="Folder"
            LibraryTitle=$LibraryTitle
            FolderName=[string]$it["FileLeafRef"]
            FolderServerRelative=$ref
            FolderUrl=Get-FullUrlFromServerRelative `
              -TenantHostUrl $TenantHostUrl `
              -ServerRelativeUrl $ref
            InheritanceStatus="Unique"
            GrantedThrough=$r.GrantedThrough
            UserLogin=$r.MemberLogin
            UserEmail=$r.MemberEmail
            Roles=$r.Roles
          }) | Out-Null
        }
      } catch {
        $SkippedOut.Add([pscustomobject]@{
          Scope="FolderACL"
          Url=$WebUrl
          Object="$LibraryTitle :: $ref"
          Reason=$_.Exception.Message
        }) | Out-Null
        continue
      }
    }
  } catch {
    $SkippedOut.Add([pscustomobject]@{
      Scope="FolderEnumeration"
      Url=$WebUrl
      Object=$LibraryTitle
      Reason=$_.Exception.Message
    }) | Out-Null
  }

  return $rows
}

# ============================= OUTPUT INIT =============================
$runId = (Get-Date).ToString("yyyyMMdd_HHmmss")
$outDir = Join-Path $OutputRoot ("Run_{0}" -f $runId)
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$transcriptPath = Join-Path $outDir ("Transcript_{0}.txt" -f $runId)
Start-Transcript -Path $transcriptPath -Force | Out-Null

$summaryOut    = New-Object System.Collections.Generic.List[object]
$resolutionOut = New-Object System.Collections.Generic.List[object]
$webOut        = New-Object System.Collections.Generic.List[object]
$libOut        = New-Object System.Collections.Generic.List[object]
$folderOut     = New-Object System.Collections.Generic.List[object]
$skippedOut    = New-Object System.Collections.Generic.List[object]

try {
  $targets = @(
  Get-Content -Path $TargetsFile -ErrorAction Stop |
  Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
  ForEach-Object { Normalize-Url $_ } |
  Where-Object { $_ } |
  Select-Object -Unique
)

  Write-QO "Targets no TXT: $($targets.Count)" "OK"
  foreach ($t in $targets) { Write-QO " - $t" "INFO" }

  $totalTargets = $targets.Count
  $targetIndex = 0
  $subwebCache = @{}

  foreach ($target in $targets) {

    $targetIndex++
    Write-Progress -Id 1 -Activity "Processando Targets" -Status "Target $targetIndex de $totalTargets" -PercentComplete (($targetIndex/$totalTargets)*100)

    Write-QO "Processando target: $target" "INFO"
    $rootUrl = Get-RootSiteUrl -TargetUrl $target
    Write-QO " RootUrl: $rootUrl" "INFO"

    if (-not $subwebCache.ContainsKey($rootUrl)) {
      $subwebCache[$rootUrl] = Get-SubWebsDeep -RootUrl $rootUrl
    }

    $allWebs = $subwebCache[$rootUrl]
    $res = Resolve-TargetToBranchWeb -TargetUrl $target -SubWebs $allWebs

    $resolutionOut.Add([pscustomobject]@{
      RunId=$runId; TargetUrl=$target; RootUrl=$rootUrl;
      Found=$res.Found; BranchWebUrl=$res.BranchWebUrl; BranchWebTitle=$res.BranchWebTitle; Note=$res.Note
    }) | Out-Null

    if (-not $res.Found) {
      $summaryOut.Add([pscustomobject]@{
        RunId=$runId; TargetUrl=$target; RootUrl=$rootUrl; BranchWebUrl="";
        WebsScanned=0; LibrariesScanned=0; FoldersScanned=0; ExternalFindings=0; Note="Target não resolvido"
      }) | Out-Null
      continue
    }

    $branchWebUrl  = $res.BranchWebUrl
    $tenantHostUrl = Get-TenantHostUrl -AnyWebUrl $branchWebUrl

    $websToScan = @(Get-WebBranchSet -AllWebs $allWebs -BranchWebUrl $branchWebUrl)
    $totalWebs = $websToScan.Count
    $webIndex = 0

    $webFindings=0; $libFindings=0; $folderFindings=0
    $libsScanned=0; $foldersScanned=0

    foreach ($w in $websToScan) {

      $webIndex++
      Write-Progress -Id 2 -ParentId 1 -Activity "Processando Sites/Subsites" -Status "$($w.Url) ($webIndex/$totalWebs)" -PercentComplete (($webIndex/$totalWebs)*100)

# --- Web ACL
try {
  Connect-QOPnP -Url $w.Url
  $web = Get-PnPWeb -Includes Title, Url, ServerRelativeUrl, HasUniqueRoleAssignments, RoleAssignments -ErrorAction Stop
  $acl = Expand-RoleAssignmentsWithSPGroupExpansion -SecurableObject $web -ScopeUrl $w.Url -SkippedOut $skippedOut

  foreach ($r in $acl.Rows) {
    if (-not $r.IsExternal) { continue }
    $webOut.Add([pscustomobject]@{
      RunId=$runId; TargetUrl=$target; RootUrl=$rootUrl; WebUrl=$w.Url;
      Level="Web"; ObjectName=$web.Title; ObjectUrl=$web.Url;
      InheritanceStatus=if($acl.HasUnique){"Unique"}else{"Inherited"};
      GrantedThrough=$r.GrantedThrough;
      UserLogin=$r.MemberLogin; UserEmail=$r.MemberEmail; Roles=$r.Roles
    }) | Out-Null
    $webFindings++
  }

  # 🔥 🔥 ADICIONE ESTE BLOCO AQUI (FALLBACK DE GRUPOS)
  $groups = Get-PnPGroup

  foreach ($g in $groups) {
    try {
      $members = Get-PnPGroupMember -Identity $g.Id

      foreach ($m in $members) {

        if (-not (Test-IsExternalPrincipal -LoginName $m.LoginName -Email $m.Email)) { continue }

        $webOut.Add([pscustomobject]@{
          RunId=$runId
          TargetUrl=$target
          RootUrl=$rootUrl
          WebUrl=$w.Url
          Level="Web"
          ObjectName=$web.Title
          ObjectUrl=$web.Url
          InheritanceStatus="Inherited"
          GrantedThrough="SPGroup: $($g.Title)"
          UserLogin=$m.LoginName
          UserEmail=$m.Email
          Roles="From Group"
        }) | Out-Null

        $webFindings++
      }

    } catch {
      $skippedOut.Add([pscustomobject]@{
        Scope="GroupDirectRead"
        Url=$w.Url
        Object=$g.Title
        Reason=$_.Exception.Message
      }) | Out-Null
    }
  }

} catch {
  $skippedOut.Add([pscustomobject]@{Scope="WebACL"; Url=$w.Url; Object=$w.Title; Reason=$_.Exception.Message}) | Out-Null
  Write-QO "Falha Web ACL: $($w.Url) => $($_.Exception.Message)" "WARN"
}

      # --- Libraries + Folder stage
      try {
        $libs = Get-LibrariesForWeb -WebUrl $w.Url
        $totalLibs = ($libs | Measure-Object).Count
        $libIndex = 0

        foreach ($lib in $libs) {
          $libsScanned++
          $libIndex++

          $pctLib = if ($totalLibs -gt 0) { ($libIndex/$totalLibs)*100 } else { 100 }
          Write-Progress -Id 3 -ParentId 2 -Activity "Processando Bibliotecas" -Status "$($lib.Title) ($libIndex/$totalLibs)" -PercentComplete $pctLib

          # Library ACL
          $aclL = Expand-RoleAssignmentsWithSPGroupExpansion -SecurableObject $lib -ScopeUrl $w.Url -SkippedOut $skippedOut
          $libRoot = ""
          try { $libRoot = $lib.RootFolder.ServerRelativeUrl } catch { $libRoot = "" }

          foreach ($r in $aclL.Rows) {
            if (-not $r.IsExternal) { continue }
            $libOut.Add([pscustomobject]@{
              RunId=$runId; TargetUrl=$target; RootUrl=$rootUrl; WebUrl=$w.Url;
              Level="Library"; ObjectName=$lib.Title;
              ObjectServerRelative=$libRoot;
              ObjectUrl=if($libRoot){ Get-FullUrlFromServerRelative -TenantHostUrl $tenantHostUrl -ServerRelativeUrl $libRoot } else { "" };
              InheritanceStatus=if($aclL.HasUnique){"Unique"}else{"Inherited"};
              GrantedThrough=$r.GrantedThrough;
              UserLogin=$r.MemberLogin; UserEmail=$r.MemberEmail; Roles=$r.Roles
            }) | Out-Null
            $libFindings++
          }

          # 🔥 FIX: HERDAR usuários externos do Web
          $webUsers = $webOut | Where-Object { $_.WebUrl -eq $w.Url }

          foreach ($wu in $webUsers) {
            $libOut.Add([pscustomobject]@{
              RunId=$runId
              TargetUrl=$target
              RootUrl=$rootUrl
              WebUrl=$w.Url
              Level="Library"
              ObjectName=$lib.Title
              ObjectServerRelative=$libRoot
              ObjectUrl=if($libRoot){ Get-FullUrlFromServerRelative -TenantHostUrl $tenantHostUrl -ServerRelativeUrl $libRoot } else { "" }
              InheritanceStatus="Inherited"
              GrantedThrough=$wu.GrantedThrough
              UserLogin=$wu.UserLogin
              UserEmail=$wu.UserEmail
              Roles=$wu.Roles
            }) | Out-Null

            $libFindings++
          }

          # Folder stage (indica etapa, sem percent real de pastas para não pesar)
          Write-Progress -Id 4 -ParentId 3 -Activity "Pastas (somente Unique ACL)" -Status "Lendo pastas em $($lib.Title)" -PercentComplete 0

          # Folder ACL (somente pastas; sem arquivos)
          $fRows = Get-FolderExternalAclRows `
          -WebUrl $w.Url `
          -LibraryTitle $lib.Title `
          -TargetUrl $target `
          -RootUrl $rootUrl `
          -TenantHostUrl $tenantHostUrl `
          -SkippedOut $skippedOut `
          -WebOut $webOut `
          -RunId $runId

          foreach ($fr in $fRows) { $folderOut.Add($fr) | Out-Null; $folderFindings++ }
          $foldersScanned += ($fRows | Measure-Object).Count

          Write-Progress -Id 4 -ParentId 3 -Activity "Pastas (somente Unique ACL)" -Completed
        }

        Write-Progress -Id 3 -ParentId 2 -Activity "Processando Bibliotecas" -Completed
      } catch {
        $skippedOut.Add([pscustomobject]@{Scope="Libraries"; Url=$w.Url; Object=$w.Title; Reason=$_.Exception.Message}) | Out-Null
        Write-QO "Falha enumerar libs: $($w.Url) => $($_.Exception.Message)" "WARN"
        Write-Progress -Id 3 -ParentId 2 -Activity "Processando Bibliotecas" -Completed
        Write-Progress -Id 4 -ParentId 3 -Activity "Pastas (somente Unique ACL)" -Completed
      }
    }

    Write-Progress -Id 2 -ParentId 1 -Activity "Processando Sites/Subsites" -Completed

    $summaryOut.Add([pscustomobject]@{
      RunId=$runId
      TargetUrl=$target
      RootUrl=$rootUrl
      BranchWebUrl=$branchWebUrl
      WebsScanned=$websToScan.Count
      LibrariesScanned=$libsScanned
      FolderFindings=$folderFindings
      ExternalFindings=($webFindings+$libFindings+$folderFindings)
      Note=""
    }) | Out-Null

    Write-QO "WebsScanned=$($websToScan.Count) | LibrariesScanned=$libsScanned | FolderFindings=$folderFindings | ExternalFindings=$($webFindings+$libFindings+$folderFindings)" "OK"
  }

  Write-Progress -Id 1 -Activity "Processando Targets" -Completed

  # Export
  $p1 = Join-Path $outDir ("01_Summary_{0}.csv" -f $runId)
  $p2 = Join-Path $outDir ("02_TargetResolution_{0}.csv" -f $runId)
  $p3 = Join-Path $outDir ("03_WebExternalAccess_{0}.csv" -f $runId)
  $p4 = Join-Path $outDir ("04_LibraryExternalAccess_{0}.csv" -f $runId)
  $p5 = Join-Path $outDir ("05_FolderExternalAccess_{0}.csv" -f $runId)
  $p6 = Join-Path $outDir ("06_Skipped_{0}.csv" -f $runId)

  Export-CsvUtf8Bom -Data $summaryOut    -Path $p1
  Export-CsvUtf8Bom -Data $resolutionOut -Path $p2
  Export-CsvUtf8Bom -Data $webOut        -Path $p3
  Export-CsvUtf8Bom -Data $libOut        -Path $p4
  Export-CsvUtf8Bom -Data $folderOut     -Path $p5
  Export-CsvUtf8Bom -Data $skippedOut    -Path $p6

  Write-QO "CSV gerados em: $outDir" "OK"

  if ($HasImportExcel) {
    $xlsx = Join-Path $outDir ("SPO_ExternalAccess_v1_{0}.xlsx" -f $runId)
    $summaryOut    | Export-Excel -Path $xlsx -WorksheetName "01_Summary" -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet
    $resolutionOut | Export-Excel -Path $xlsx -WorksheetName "02_TargetResolution" -AutoSize -FreezeTopRow -BoldTopRow -Append
    $webOut        | Export-Excel -Path $xlsx -WorksheetName "03_WebExternalAccess" -AutoSize -FreezeTopRow -BoldTopRow -Append
    $libOut        | Export-Excel -Path $xlsx -WorksheetName "04_LibraryExternalAccess" -AutoSize -FreezeTopRow -BoldTopRow -Append
    $folderOut     | Export-Excel -Path $xlsx -WorksheetName "05_FolderExternalAccess" -AutoSize -FreezeTopRow -BoldTopRow -Append
    $skippedOut    | Export-Excel -Path $xlsx -WorksheetName "06_Skipped" -AutoSize -FreezeTopRow -BoldTopRow -Append
    Write-QO "XLSX gerado: $xlsx" "OK"
  } else {
    Write-QO "ImportExcel não encontrado. XLSX não gerado." "WARN"
  }

  Write-QO "Concluído. Transcript: $transcriptPath" "OK"

} finally {
  Stop-Transcript | Out-Null
  Write-Progress -Id 4 -Completed
  Write-Progress -Id 3 -Completed
  Write-Progress -Id 2 -Completed
  Write-Progress -Id 1 -Completed
}
