param(
  [Parameter()]
  [string]$tenant
)

function ShowProgressBar {
  param(
    [Parameter(Mandatory)]
    $Current,
    [Parameter(Mandatory)]
    $Total
  )

  Write-Progress -Activity "Buscando usuários locais..." `
    -Status "Aguarde: $($Current) de $($Total) ($([Math]::Round(($Current / $Total) * 100))%)" `
    -PercentComplete (($Current / $Total) * 100)
  Start-Sleep -Milliseconds 50
}

function NormalizeFileName {
  param([string]$str)

  $normalized = $str.Normalize([System.Text.NormalizationForm]::FormD)
  $asciiOnly = -join ($normalized.ToCharArray() | Where-Object {
      [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark'
    })

  return $asciiOnly.ToLower() -replace '[^a-z0-9]+', '_'
}

function OpenNewTenantConnection {
  Disconnect-MgGraph | Out-Null
  Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
  Write-Host "🔄 Conectando ao Microsoft Entra ID..." -ForegroundColor Yellow
}

function GetLocalUsers {
  Write-Host "🔍 Buscando todos os usuários..."
  $allUsers = Get-MgUser -All -Property "displayName,userPrincipalName,UserType"

  $localUsers = $allUsers | Where-Object { $_.UserPrincipalName -notmatch '#EXT#' }

  $totalUsers = $localUsers.Count
  $progressUsers = @()

  for ($i = 0; $i -lt $totalUsers; $i++) {
    ShowProgressBar -Current ($i + 1) -Total $totalUsers
    $progressUsers += [PSCustomObject]@{
      DisplayName       = $localUsers[$i].DisplayName
      UserPrincipalName = $localUsers[$i].UserPrincipalName
    }
  }

  return $progressUsers
}

function Main {
  Clear-Host
  Write-Host ""
  Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
  Write-Host "║              🪟 Microsoft Entra ID - Usuários por licença          ║" -ForegroundColor Cyan
  Write-Host "║--------------------------------------------------------------------║" -ForegroundColor Cyan
  Write-Host "║ Autor      : Natan Felipe de Oliveira                              ║" -ForegroundColor Cyan
  Write-Host "║ Descrição  : Mostra e exporta os usuários locais do Tenant         ║" -ForegroundColor Cyan
  Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
  Write-Host ""

  if (![string]::IsNullOrEmpty($tenant)) {
    OpenNewTenantConnection
  }

  $context = Get-MgContext
  if (!$context.Account) {
    Write-Host "❌ Não foi possível se conectar ao tenant" -ForegroundColor Red
    exit 1
  }

  Write-Host "🔄 Conectado ao tenant: $((Get-MgOrganization).DisplayName)" -ForegroundColor Yellow

  $localUsers = GetLocalUsers

  Write-Host ""
  Write-Host "📊 Total de usuários locais encontrados: $($localUsers.Count)" -ForegroundColor Green
 
  $sortedLocalUsers = $localUsers | ForEach-Object -Begin { $i = 1 } -Process {
    [PSCustomObject]@{
      Rank       = $i
      Nome       = $_.DisplayName
      Email      = $_.UserPrincipalName
    }
    $i++
  }
  $sortedLocalUsers | Format-Table -AutoSize

  # Exporta CSV
  $dateStr = (Get-Date).ToString("ddMMyyyy_HHmmss")
  $csvDir = Join-Path $PSScriptRoot "..\output"

  if (-not (Test-Path $csvDir)) {
    New-Item -Path $csvDir -ItemType Directory | Out-Null
  }

  $orgName = (Get-MgOrganization).DisplayName
  $sanitizedOrgName = NormalizeFileName $orgName
  $csvPath = Join-Path $csvDir "${sanitizedOrgName}_${dateStr}_usuarios_locais.csv"

  $localUsers | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

  Write-Host "💾 Lista exportada para: $csvPath" -ForegroundColor Yellow
}

try {
  $start = Get-Date
  Main -tenant $tenant
  $end = Get-Date
  $duration = $end - $start
  Write-Host "`n⏱️ Tempo total: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
}
catch {
  Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor DarkRed
  exit 1
}
