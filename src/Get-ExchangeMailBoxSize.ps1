# - Acesse a documentação do script através do link abaixo 👇
# - https://github.com/natanzeraa/scripts-and-automation/blob/main/README/PowerShell/GetExchangeMailBoxSize.md

Clear-Host
Write-Host "`nIniciando contagem de caixas de e-mail..." -ForegroundColor Gray

$mailboxes = Get-Mailbox -ResultSize Unlimited
$mailboxesCount = $mailboxes.Count

if ($mailboxesCount -eq 0) {
    Write-Host "Nenhuma caixa de e-mail encontrada." -ForegroundColor Red
    exit
}

$orgName = (Get-OrganizationConfig).DisplayName
Write-Host "`nOrganização: $orgName" -ForegroundColor Gray
Write-Host "`nTotal de caixas de e-mail: $mailboxesCount" -ForegroundColor Gray

[int]$topRankingCount = Read-Host "`nQuantas caixas de e-mail mais ocupadas você deseja visualizar no ranking"

function Show-Progress($current, $total) {
    Write-Progress -Activity "Coletando caixas de e-mail" `
        -Status "$current de $total processado(s) ($([math]::Round(($current / $total) * 100))%)" `
        -PercentComplete (($current / $total) * 100)
}

function Measure-AllMailboxesSize {
    param (
        [Parameter(Mandatory)]
        $mailboxesData
    )

    $totalBytes = ($mailboxesData | Measure-Object -Property ByteSize -Sum).Sum
    $totalGB = [math]::Round($totalBytes / 1GB, 2)
    $totalTB = [math]::Round($totalBytes / 1TB, 2)

    Write-Host "`nUso total de todas as $mailboxesCount caixas de e-mail: $totalBytes bytes (~$totalGB GB) (~$totalTB TB)" -ForegroundColor DarkYellow
}

function Measure-RankedMailboxesSize {
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

function Measure-MailboxesSizeMean {
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

function Convert-StringToBytes {
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

function Get-MailboxUsageReport {
    $results = @()
    $current = 0
    $errors = @()
    $startTime = Get-Date

    Write-Host "`nAguarde... coletando estatísticas...`n"

    foreach ($mailbox in $mailboxes) {
        $current++
        Show-Progress -current $current -total $mailboxesCount

        try {
            $stats = Get-MailboxStatistics -Identity $mailbox.Guid -ErrorAction Stop
            
            if ($stats.TotalItemSize -and $stats.TotalItemSize.Value) {
                $rawSize = $stats.TotalItemSize.ToString()
                
                $bytes = if ($rawSize -match '\(([\d,]+)\sbytes\)') { [int64]($matches[1] -replace ',', '') } else { 0 }
                
                $prohibitSendQuota = $mailbox.ProhibitSendQuota
                $quotaBytes = Convert-StringToBytes -prohibitSendQuota $prohibitSendQuota
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

    Write-Progress -Activity "Coletando caixas de e-mail" -Completed

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

    Measure-RankedMailboxesSize -mailboxesData $topMailboxes -rankingCount $topRankingCount
    Measure-AllMailboxesSize -mailboxesData $results
    Measure-MailboxesSizeMean -mailboxesData $results -totalMailboxesCount $mailboxesCount

    $csvDir = Join-Path $PSScriptRoot "..\output"
    if (-not (Test-Path $csvDir)) {
        New-Item -Path $csvDir -ItemType Directory | Out-Null
    }

    $csvPath = Join-Path $csvDir "top_${topRankingCount}_caixas_de_email.csv"

    $topMailboxes | Select-Object Name, Email, Size, Quota, Usage, Free, Enviados | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8
    Write-Host "`nResultado exportado para: $csvPath" -ForegroundColor Green

    $duration = (Get-Date) - $startTime
    Write-Host "`nDuração: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s`n" -ForegroundColor DarkCyan

    if ($errors.Count -gt 0) {
        Write-Host "`n[ERRO] Erros durante a coleta:" -ForegroundColor Red
        $errors | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    }
}

Get-MailboxUsageReport
