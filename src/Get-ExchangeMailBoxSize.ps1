# - Acesse a documentação do script através do link abaixo 👇
# - https://github.com/natanzeraa/scripts-and-automation/blob/main/README/PowerShell/GetExchangeMailBoxSize.md

param(
    [Parameter()]
    [string]$tenant
)

function OpenNewTenantConnection {
    try {
        Write-Host "`n🔍 Verificando sessões existentes..." -ForegroundColor Yellow
        $graphContext = Get-MgContext

        if ($graphContext.Account) {
            Write-Host "Encerrando sessão do Microsoft Graph: $($graphContext.Account)"
            Disconnect-MgGraph | Out-Null
        }
    }
    catch {
        Write-Host "Não foi possível verificar sessão do Graph." -ForegroundColor DarkRed
        exit 1
    }

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "Não foi possível encerrar sessão do Exchange Online." -ForegroundColor DarkRed
        exit 1
    }

    Write-Host "`n🔐 Autenticação necessária!" -ForegroundColor Cyan
    Write-Host "Será aberta uma URL para você autenticar usando um código de dispositivo."
    Write-Host "Caso não apareça automaticamente, acesse https://microsoft.com/devicelogin manualmente."
    Write-Host ""

    try {
        Connect-ExchangeOnline -UserPrincipalName $tenant -ShowBanner:$false -Device
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
        Write-Host "🔄 Conectado ao Exchange Online com: $tenant" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Erro ao conectar com o Exchange: $($_.Exception.Message)" -ForegroundColor DarkRed
        exit 1
    }
}

function CheckExistentContext {
    $context = Get-MgContext
    
    if (!$context.Account) {
        Write-Host "❌ Não foi possível se conectar ao tenant" -ForegroundColor DarkRed
        exit 1
    }

    return $context
}

function CountMailBoxes {
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    $mailboxesCount = $mailboxes.Count

    if ($mailboxesCount -eq 0) {
        Write-Host "Nenhuma caixa de e-mail encontrada." -ForegroundColor DarkRed
        exit 1
    }

    Write-Host "`nTotal de caixas de e-mail: $mailboxesCount"

    return $mailboxes, $mailboxesCount
}

function ShowProgress($current, $total) {
    Write-Progress -Activity "Coletando caixas de e-mail" `
        -Status "$current de $total processado(s) ($([math]::Round(($current / $total) * 100))%)" `
        -PercentComplete (($current / $total) * 100)
}

function MeasureAllMailboxesSize {
    param (
        [Parameter(Mandatory)]
        $mailboxesData
    )

    $totalBytes = ($mailboxesData | Measure-Object -Property ByteSize -Sum).Sum
    $totalGB = [math]::Round($totalBytes / 1GB, 2)
    $totalTB = [math]::Round($totalBytes / 1TB, 2)

    Write-Host "`nUso total de todas as $mailboxesCount caixas de e-mail: $totalBytes bytes (~$totalGB GB) (~$totalTB TB)" -ForegroundColor DarkYellow
}

function MeasureRankedMailboxesSize {
    param (
        [Parameter(Mandatory)]
        $mailboxesData,

        [Parameter(Mandatory)]
        [int]$rankingCount
    )

    $totalBytes = ($mailboxesData | Measure-Object -Property ByteSize -Sum).Sum
    $totalGB = [math]::Round($totalBytes / 1GB, 2)
    $totalTB = [math]::Round($totalBytes / 1TB, 2)

    Write-Host "`nUso total das $rankingCount maiores caixas de e-mail: $totalBytes bytes (~$totalGB GB) (~$totalTB TB)" -ForegroundColor DarkYellow
}

function MeasureMailboxesSizeMean {
    param (
        [Parameter(Mandatory)]
        $mailboxesData,

        [Parameter(Mandatory)]
        $totalMailboxesCount
    )

    $totalBytes = ($mailboxesData | Measure-Object -Property ByteSize -Sum).Sum
    $mailboxesMedian = ($totalBytes / $totalMailboxesCount)
    $totalGB = [math]::Round($mailboxesMedian / 1GB, 2)
    $totalTB = [math]::Round($mailboxesMedian / 1TB, 4)

    Write-Host "`nUso médio por caixa de e-mail: $mailboxesMedian bytes (~$totalGB GB) (~$totalTB TB)" -ForegroundColor DarkYellow
}

function ConvertStringToBytes {
    param (
        [Parameter(Mandatory)]
        [string]$prohibitSendQuota
    )

    if ($prohibitSendQuota -eq "Unlimited") {
        return 0
    }

    $normalized = $prohibitSendQuota.Trim().ToUpper()

    if ($normalized -match '\(([\d,]+)\sbytes\)') { 
        $bytes = [int64]$matches[1] -replace ',', ''
        return $bytes
    }
    else {
        return 0
    }
}

function GetMailboxUsageReport {
    param(
        [Parameter(Mandatory)]
        [int64]$ranking,

        [Parameter(Mandatory)]
        $mailboxes,

        [Parameter(Mandatory)]
        [int]$mailboxesCount
    )

    $results = @()
    $current = 0
    $errors = @()

    Write-Host "`nAguarde... coletando estatísticas...`n"

    foreach ($mailbox in $mailboxes) {
        $current++
        ShowProgress -current $current -total $mailboxesCount

        try {
            $stats = Get-MailboxStatistics -Identity $mailbox.Guid -ErrorAction Stop
            
            if ($stats.TotalItemSize -and $stats.TotalItemSize.Value) {
                $rawSize = $stats.TotalItemSize.ToString()
                
                $bytes = if ($rawSize -match '\(([\d,]+)\sbytes\)') { [int64]($matches[1] -replace ',', '') } else { 0 }
                
                $prohibitSendQuota = $mailbox.ProhibitSendQuota
                $quotaBytes = ConvertStringToBytes -prohibitSendQuota $prohibitSendQuota
                $percentUsed = if ($quotaBytes -gt 0) { [math]::Round(($bytes / $quotaBytes) * 100, 2) } else { 0 }
                $freeGB = if ($quotaBytes -gt 0) { [math]::Round(($quotaBytes - $bytes) / 1GB, 2) } else { 0 } 

                $results += [PSCustomObject]@{
                    Name     = $mailbox.DisplayName
                    Email    = $mailbox.UserPrincipalName
                    Size     = $stats.TotalItemSize
                    ByteSize = $bytes
                    Sent     = $stats.ItemCount
                    Quota    = if ($quotaBytes -gt 0) { [math]::Round($quotaBytes / 1GB, 2) } else { 'Ilimitado' }
                    Usage    = $percentUsed
                    Free     = $freeGB
                }
            }
            else {
                $errors += "$($mailbox.DisplayName) <$($mailbox.UserPrincipalName)>"
            }
        }
        catch {
            $errors += "$($mailbox.DisplayName) <$($mailbox.UserPrincipalName)>: $($_.Exception.Message)"
        }
    }

    return $results, $errors
}

function NormalizeFileName {
    param([string]$str)

    $normalized = $str.Normalize([System.Text.NormalizationForm]::FormD)
    $asciiOnly = -join ($normalized.ToCharArray() | Where-Object {
            [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark'
        })

    return $asciiOnly.ToLower() -replace '[^a-z0-9]+', '_'
}

function CreateAndSaveCsvToOrgFolder {
    $orgName = (Get-MgOrganization).DisplayName

    $normalizedOrgName = NormalizeFileName -str $orgName

    $folderPath = Join-Path $PSScriptRoot "..\$normalizedOrgName"
    
    if (-not (Test-Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
    }
    
    $path = Join-Path $folderPath "${normalizedOrgName}_top_${topRankingCount}_caixas_de_email.csv"
    return $path
}

function Main {
    Clear-Host
    Write-Host ""
    Write-Host "╔═════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║              🪟 Microsoft Entra ID - Ranking de caixas de email          ║" -ForegroundColor Cyan
    Write-Host "║-------------------------------------------------------------------------║" -ForegroundColor Cyan
    Write-Host "║ Autor      : Natan Felipe de Oliveira                                   ║" -ForegroundColor Cyan
    Write-Host "║ Descrição  : Exibe um ranking de caixas de email ordenadas por tamanho  ║" -ForegroundColor Cyan
    Write-Host "╚═════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    if (![string]::IsNullOrEmpty($tenant)) {
        OpenNewTenantConnection
    }

    $context = CheckExistentContext
    
    Write-Host "Sessão iniciada em: $(($context).Account)"
    Write-Host "`nConectado ao tenant: $((Get-MgOrganization).DisplayName)"
    Write-Host "`nIniciando contagem de caixas de e-mail..."

    $mailboxes, $mailboxesCount = CountMailBoxes

    [int]$topRankingCount = Read-Host "`nQuantas caixas de e-mail mais ocupadas você deseja visualizar no ranking"

    $results, $errors = GetMailboxUsageReport -ranking $topRankingCount -mailboxes $mailboxes -mailboxesCount $mailboxesCount

    $topMailboxes = $results | Sort-Object -Property ByteSize -Descending | Select-Object -First $topRankingCount

    $rankedTopMailboxes = $topMailboxes | ForEach-Object -Begin { $i = 1 } -Process {
        [PSCustomObject]@{
            Rank       = $i
            Nome       = $_.Name
            Email      = $_.Email
            Usado      = $_.Size
            Capacidade = "$($_.Quota) GB"
            "Uso (%)"  = "$($_.Usage)%"
            Disponivel = "$($_.Free) GB"
            Enviados   = $_.Sent
        }
        $i++
    }
    
    Write-Host "`nTop $topRankingCount caixas de e-mail mais ocupadas:`n" -ForegroundColor Yellow
    $rankedTopMailboxes | Format-Table -AutoSize

    MeasureRankedMailboxesSize -mailboxesData $topMailboxes -rankingCount $topRankingCount
    MeasureAllMailboxesSize -mailboxesData $results
    MeasureMailboxesSizeMean -mailboxesData $results -totalMailboxesCount $mailboxesCount
    
    $path = CreateAndSaveCsvToOrgFolder
    $topMailboxes | Select-Object Name, Email, Size, Quota, Usage, Free, Sent | Export-Csv -Path $path -NoTypeInformation -Encoding utf8
    Write-Host "`nResultado exportado para: $path" -ForegroundColor Green

    if ($errors.Count -gt 0) {
        Write-Host "`n[ERRO] Erros durante a coleta:" -ForegroundColor DarkRed
        $errors | ForEach-Object { Write-Host $_ -ForegroundColor DarkRed }
    }
}

try {
    $start = Get-Date
    Main -tenant $tenant
    $time = (Get-Date) - $start
    Write-Host ("Tempo: {0:hh\:mm\:ss}" -f $time)
}
catch {
    Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor DarkRed
    exit 1
}
