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
        $licenseMatch = $licenses | Where-Object { ($_.SkuId.ToString()).ToLower() -eq ($assigned.skuId.ToString()).ToLower() }
        if ($licenseMatch) {
          $userLicenses += $licenseMatch.SkuPartNumber
        }
      }
    }

    if ($userLicenses.Count -gt 0) {
      $withLicenses += [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Licenses          = $userLicenses -join ", "
      }
    }
    else {
      $withoutLicenses += [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Licenses          = "UNLICENSED"
      }
    }
  }

  return $withLicenses, $withoutLicenses
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
  Write-Host "🔄 Conectando ao tenant $((Get-MgOrganization).DisplayName)..."
  Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

  $usersWithLicenses, $usersWithoutLicenses = GetUsersWithLicenses

  if ($usersWithLicenses.Count -eq 0) {
    Write-Host "Nenhum usuário com licença encontrado." -ForegroundColor Yellow
  }
  else {
    Write-Host "🔎 Usuários encontrados com licenças:"
    $usersWithLicenses | Sort-Object DisplayName | Format-Table -AutoSize
  }

  if ($usersWithoutLicenses.Count -eq 0) {
    Write-Host "Todos os usuários possuem licença." -ForegroundColor Green
  }
  else {
    Write-Host "`n🔎 Usuários sem licença:"
    $usersWithoutLicenses | Sort-Object DisplayName | Format-Table -AutoSize
  }

  $dateObj = Get-Date
  $dateStr = $dateObj.ToString("ddMMyyyy_HHmmss")
  $csvDir = Join-Path $PSScriptRoot "..\output"
    
  if (-not (Test-Path $csvDir)) { 
    New-Item -Path $csvDir -ItemType Directory | Out-Null 
  }

  $usersWithLicensesCsvPath = Join-Path $csvDir "$($dateStr)_users_with_license.csv"
  $usersWithoutLicensesCsvPath = Join-Path $csvDir "$($dateStr)_users_without_license.csv"

  $usersWithLicenses | Export-Csv -Path $usersWithLicensesCsvPath -NoTypeInformation -Encoding utf8
  $usersWithoutLicenses | Export-Csv -Path $usersWithoutLicensesCsvPath -NoTypeInformation -Encoding utf8
  
  Write-Host "📊 Total de usuários com licenças: $($usersWithLicenses.Count)" -ForegroundColor Green
  Write-Host "🔄 Exportando usuários com licença para CSV: $usersWithLicensesCsvPath" -ForegroundColor Yellow
  
  Write-Host "`n📊 Total de usuários sem licenças: $($usersWithoutLicenses.Count)" -ForegroundColor Red
  Write-Host "🔄 Exportando usuários com licença para CSV: $usersWithoutLicensesCsvPath" -ForegroundColor Yellow
}

try {
  $start = Get-Date
  Main
  $end = Get-Date
  $time = $end - $start
  Write-Host "`n⏱️ Tempo total: $($time.Hours)h $($time.Minutes)m $($time.Seconds)s"
}
catch {
  Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor DarkRed
  Write-Host "❌ Detalhes: $($_.ErrorDetails.RecommendedAction)" -ForegroundColor DarkRed
  exit 1
}
