function Get-ValidDate($prompt) {
    while ($true) {
        $userInput = Read-Host $prompt
        try {
            return [datetime]::ParseExact($userInput, 'dd-MM-yyyy', $null)
        }
        catch {
            Write-Host "❌ Formato inválido. Use o formato dd-MM-yyyy." -ForegroundColor Red
        }
    }
}

function Show-Progress($current, $total) {
    Write-Progress -Activity "Buscando usuários" `
        -Status "$current de $total processado(s) ($([math]::Round(($current / $total) * 100))%)" `
        -PercentComplete (($current / $total) * 100)
}

function Get-UsersWithCreationDate {
    param(
        [Parameter(Mandatory)]
        $users,
        [Parameter(Mandatory)]
        $newUsersCount,
        [Parameter(Mandatory)]
        $guestOrMemberPreference
    )

    $results = @()
    $i = 0
    
    foreach ($item in $users) {
        $userMail = $item.Mail
        $i++
        
        Show-Progress -current $i -total $newUsersCount
        
        if ([string]::IsNullOrWhiteSpace($userMail)) {
            Write-Host "⚠️ Usuário '$($user.DisplayName)' não possui e-mail válido. Ignorando." -ForegroundColor DarkYellow
        }

        try {
            $filter = ""

            if ($guestOrMemberPreference.ToLower() -eq "guest" -or $guestOrMemberPreference.ToLower() -eq "member") {
                $filter += "mail eq '$userMail' and userType eq '$guestOrMemberPreference'"
            }

            if ([string]::IsNullOrWhiteSpace($guestOrMemberPreference)) {
                $filter += "mail eq '$userMail'"
            }
 
            $fullUser = Get-MgUser -Filter $filter -Property DisplayName, Mail, UserPrincipalName, CreatedDateTime, LastPasswordChangeDateTime, AccountEnabled, UserType

            if ($fullUser) {
                $results += [PSCustomObject]@{
                    Name                       = $fullUser.DisplayName
                    Mail                       = $fullUSer.Mail
                    UPN                        = $fullUser.UserPrincipalName
                    CreatedDateTime            = $fullUser.CreatedDateTime
                    LastPasswordChangeDateTime = $fullUser.LastPasswordChangeDateTime.ToLocalTime()
                    AccountEnabled             = $fullUser.AccountEnabled ? "Sim" : "Não"
                    UserType                   = if ($fullUser.UserType -eq "Guest") { "Convidado" } else { "Membro" }
                }
            }
        }
        catch {
            Write-Host "❌ Erro ao buscar usuário com e-mail '$userMail': $_" -ForegroundColor Yellow
        }
    } Write-Host "`nUsuários criados/convidados entre $startDateObj e $endDateObj`n" -ForeGroundColor Green

    return $results
}


function Main {
    Clear-Host
    Write-Host ""
    Write-Host "╔═════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║              🪟 Microsoft Entra ID - Usuários criados       ║" -ForegroundColor Cyan
    Write-Host "║-------------------------------------------------------------║" -ForegroundColor Cyan
    Write-Host "║ Autor      : Natan Felipe de Oliveira                       ║" -ForegroundColor Cyan
    Write-Host "║ Descrição  : Filtra usuários Entra ID com base no intervalo ║" -ForegroundColor Cyan
    Write-Host "║              de datas da criação da conta.                  ║" -ForegroundColor Cyan
    Write-Host "╚═════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "🔄 Conectando ao Microsoft Entra ID..."
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
    Write-Host "✅ Conectado com sucesso!`n"

    $start = Get-Date

    $startDateObj = Get-ValidDate "📅 Insira a data de início (dd-MM-yyyy)"

    $endDateObj = Get-ValidDate "📅 Insira a data final (dd-MM-yyyy)"

    $guestOrMemberPreference = Read-Host "🤔 Para exibir Convidados (Guest) | Membros (Member) | Ambos (Pressione Enter) "

    $startDate = $startDateObj.ToString("yyyy-MM-ddTHH:mm:ssZ")

    $endDate = $endDateObj.ToString("yyyy-MM-ddTHH:mm:ssZ")

    Write-Host "`n🔍 Buscando usuários criados entre: $(($startDateObj).ToString("dd/MM/yyyy")) e $(($endDateObj).ToString("dd/MM/yyyy"))`n"

    $filter = "createdDateTime ge $startDate and createdDateTime le $endDate"

    $newUsers = Get-MgUser -Filter $filter -All

    $newUsersCount = $newUsers.Count

    $outputObj = Get-UsersWithCreationDate -users $newUsers -guestOrMemberPreference $guestOrMemberPreference -newUsersCount $newUsersCount

    $table = $outputObj | ForEach-Object -Begin { $i = 1 } -Process {
        [PSCustomObject]@{
            "#"               = $i
            "Nome"            = $_.Name
            "Email"           = $_.Mail
            "UPN"             = $_.UPN
            "Data de criação" = $_.CreatedDateTime
            "Senha alterada"  = $_.LastPasswordChangeDateTime
            "Conta ativa"     = $_.AccountEnabled
            "Tipo do usuário" = $_.UserType
        }
        $i++
    }

    $table | Format-Table -AutoSize

    $end = Get-Date

    $time = $end - $start

    Write-Host "Tempo: $($time.Hours):$($time.Minutes):$($time.Seconds)"
}

Main
