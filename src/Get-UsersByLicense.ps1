function ShowProgressBar {
  param(
    [Parameter(Mandatory)]
    $Current,
    [Parameter(Mandatory)]
    $Total
  )

  Write-Progress -Activity "Buscando logins mais recentes..." `
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

function GetUsersWithLicenses {
  Write-Host "🔍 Calculando total de usuários..."
  $users = Get-MgUser -All
  $userCount = $users.Count
  Write-Host "🔍 Usuários na base $($userCount)..."

  Write-Host "🔍 Buscando licenças disponíveis..."
  $licenses = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId

  if (!$licenses) {
    Write-Host "❌ Nenhuma licença encontrada." -ForegroundColor Red
    return @()
  }

  Write-Host "🔍 Buscando todos os usuários (isso pode demorar dependendo da quantidade)..."
  $allUsers = Get-MgUser -All -Property "displayName,userPrincipalName,assignedLicenses"

  $totalUsers = $allUsers.Count
  $withLicenses = @()
  $withoutLicenses = @()

  for ($i = 0; $i -lt $totalUsers; $i++) {
    $user = $allUsers[$i]

    ShowProgressBar -Current ($i + 1) -Total $totalUsers

    $userLicenses = @()
    if ($user.assignedLicenses) {
      foreach ($assigned in $user.assignedLicenses) {
        $licenseMatch = $licenses | Where-Object {
          ($_.SkuId.ToString()).ToLower() -eq ($assigned.skuId.ToString()).ToLower()
        }
        if ($licenseMatch) {
          $userLicenses += $licenseMatch.SkuPartNumber
        }
      }
    }

    $userObject = [PSCustomObject]@{
      DisplayName       = $user.DisplayName
      UserPrincipalName = $user.UserPrincipalName
      Licenses          = if ($userLicenses.Count -gt 0) { $userLicenses -join ", " } else { "UNLICENSED" }
    }

    if ($userLicenses.Count -gt 0) {
      $withLicenses += $userObject
    }
    else {
      $withoutLicenses += $userObject
    }
  }

  return $withLicenses, $withoutLicenses
}

function VerifyTenantConnection {
  $mgContext = Get-MgContext
  $isConnected = $mgContext.Account
  $choice = ""

  if ($isConnected) {
    Write-Host "🔄 Já conectado ao tenant: $((Get-MgOrganization).DisplayName)" -ForegroundColor Yellow
    $choice += Read-Host "Deseja trocar de tenant? (S/N)"
  }

  if (-not $isConnected) {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
    Write-Host "🔄 Conectando ao Microsoft Entra ID..." -ForegroundColor Yellow
  }

  if ($choice.Trim().ToLower() -eq "s") {
    Disconnect-MgGraph | Out-Null
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
    Write-Host "✅ Conectado ao tenant: $((Get-MgOrganization).DisplayName)" -ForegroundColor Green
  }
}

function Main {
  Clear-Host
  Write-Host ""
  Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
  Write-Host "║              🪟 Microsoft Entra ID - Usuários por licença          ║" -ForegroundColor Cyan
  Write-Host "║--------------------------------------------------------------------║" -ForegroundColor Cyan
  Write-Host "║ Autor      : Natan Felipe de Oliveira                              ║" -ForegroundColor Cyan
  Write-Host "║ Descrição  : Mostra os usuários com base na licença requerida      ║" -ForegroundColor Cyan
  Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
  Write-Host ""

  VerifyTenantConnection
  
  $usersWithLicenses, $usersWithoutLicenses = GetUsersWithLicenses

  if ($usersWithLicenses.Count -eq 0) {
    Write-Host "Nenhum usuário com licença encontrado." -ForegroundColor Yellow
  }
  else {
    Write-Host "🔎 Usuários com licenças:"
    $usersWithLicenses | Sort-Object DisplayName | Format-Table -AutoSize
  }

  if ($usersWithoutLicenses.Count -eq 0) {
    Write-Host "Todos os usuários possuem licença." -ForegroundColor Green
  }
  else {
    Write-Host "`n🔎 Usuários sem licença:"
    $usersWithoutLicenses | Sort-Object DisplayName | Format-Table -AutoSize
  }

  $dateStr = (Get-Date).ToString("ddMMyyyy_HHmmss")
  $csvDir = Join-Path $PSScriptRoot "..\output"

  if (-not (Test-Path $csvDir)) {
    New-Item -Path $csvDir -ItemType Directory | Out-Null
  }

  $orgName = (Get-MgOrganization).DisplayName
  $sanitizedOrgName = NormalizeFileName $orgName

  $usersWithLicensesCsvPath = Join-Path $csvDir "${sanitizedOrgName}_${dateStr}_users_with_license.csv"
  $usersWithoutLicensesCsvPath = Join-Path $csvDir "${sanitizedOrgName}_${dateStr}_users_without_license.csv"

  $usersWithLicenses | Export-Csv -Path $usersWithLicensesCsvPath -NoTypeInformation -Encoding utf8
  $usersWithoutLicenses | Export-Csv -Path $usersWithoutLicensesCsvPath -NoTypeInformation -Encoding utf8

  Write-Host "`n📊 Total com licença: $($usersWithLicenses.Count)" -ForegroundColor Green
  Write-Host "💾 Exportado: $usersWithLicensesCsvPath"

  Write-Host "`n📊 Total sem licença: $($usersWithoutLicenses.Count)" -ForegroundColor Red
  Write-Host "💾 Exportado: $usersWithoutLicensesCsvPath"
}

try {
  $start = Get-Date
  Main
  $end = Get-Date
  $duration = $end - $start
  Write-Host "`n⏱️ Tempo total: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
}
catch {
  Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor DarkRed
  Write-Host "❌ Detalhes: $($_.ErrorDetails.RecommendedAction)" -ForegroundColor DarkRed
  exit 1
}
