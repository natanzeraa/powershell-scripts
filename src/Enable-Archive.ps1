function Main {
  Clear-Host
  Write-Host ""
  Write-Host "╔══════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
  Write-Host "║ Script de Automação para Exchange Online                                         ║" -ForegroundColor Cyan
  Write-Host "║----------------------------------------------------------------------------------║" -ForegroundColor Cyan
  Write-Host "║ Função    : Habilita Arquivo Morto Expansivo (Auto-Expanding Archive) e ativa MRM║" -ForegroundColor Cyan
  Write-Host "║ Objetivo  : Otimizar espaço em mailboxes, limpando itens antigos e aplicando MRM ║" -ForegroundColor Cyan
  Write-Host "║ Recursos  : - Verificação de licenças (E3)                                       ║" -ForegroundColor Cyan
  Write-Host "║             - Habilitação de Arquivo Morto Expansivo                             ║" -ForegroundColor Cyan
  Write-Host "║             - Ativação do Managed Folder Assistant (MRM)                         ║" -ForegroundColor Cyan
  Write-Host "║             - Execução contínua em loop com intervalos de 10 minutos             ║" -ForegroundColor Cyan
  Write-Host "║ Compatível: PowerShell 5.1 e 7+                                                  ║" -ForegroundColor Cyan
  Write-Host "║ Autor     : Natan Felipe de Oliveira                                             ║" -ForegroundColor Cyan
  Write-Host "╚══════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
  Write-Host ""

  # 1. Ajustar política de execução apenas para esta sessão
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
  Write-Host "Política de execução ajustada temporariamente para RemoteSigned." -ForegroundColor Green

  # 2. Garantir que módulos estão instalados
  if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    Write-Host "Instalando módulo ExchangeOnlineManagement..." -ForegroundColor Yellow
  }
  if (-not (Get-Module -ListAvailable -Name MSOnline)) {
    Install-Module -Name MSOnline -Scope CurrentUser -Force
    Write-Host "Instalando módulo MSOnline..." -ForegroundColor Yellow
  }

  Import-Module ExchangeOnlineManagement
  Import-Module MSOnline

  # 3. Lista de emails a serem percorridos
  $mailboxList = @(
    "export11@bellofoods.com.br",
    "export27@bellofoods.com.br",
    "export7@bellofoods.com.br",
    "ana.reis@belloalimentos.com.br",
    "export14@bellofoods.com.br",
    "export40@bellofoods.com.br",
    "export15@bellofoods.com.br",
    "export41@bellofoods.com.br",
    "export8@bellofoods.com.br",
    "export25@bellofoods.com.br",
    "export4@bellofoods.com.br",
    "export32@bellofoods.com.br",
    "export9@bellofoods.com.br",
    "erik.kruger@belloalimentos.com.br",
    "export@coprimar.com.br",
    "exportacao@belloalimentos.com.br",
    "exportcvel@belloalimentos.com.br",
    "export@bellofoods.com.br",
    "export29@bellofoods.com.br",
    "export20@bellofoods.com.br",
    "export38@bellofoods.com.br",
    "export18@bellofoods.com.br",
    "export39@bellofoods.com.br",
    "export17@bellofoods.com.br",
    "export2@bellofoods.com.br",
    "export16@bellofoods.com.br",
    "export26@bellofoods.com.br",
    "export28@bellofoods.com.br",
    "noujaim@bellofoods.com.br",
    "export5@bellofoods.com.br",
    "export6@bellofoods.com.br",
    "export22@bellofoods.com.br",
    "karoline.pinzl@belloalimentos.com.br",
    "export35@bellofoods.com.br",
    "export30@bellofoods.com.br",
    "export34@bellofoods.com.br",
    "lucas.lima@bellofoods.com.br",
    "export21@bellofoods.com.br",
    "export33@bellofoods.com.br",
    "export3@bellofoods.com.br",
    "paula.stein@bellofoods.com.br",
    "export23@bellofoods.com.br",
    "export36@bellofoods.com.br",
    "export31@bellofoods.com.br",
    "export37@bellofoods.com.br",
    "export10@bellofoods.com.br",
    "export13@bellofoods.com.br",
    "export1@bellofoods.com.br",
    "export12@bellofoods.com.br",
    "export24@bellofoods.com.br",
    "export19@bellofoods.com.br"
  )

  $UserUPN = "seg.info@belloalimentos.com.br"

  while ($true) {
    try {
      # 4. Conectar ao Exchange Online
      Connect-MsolService
      Write-Host "Conectando ao Exchange Online como $UserUPN..." -ForegroundColor Yellow
      Connect-ExchangeOnline -UserPrincipalName $UserUPN -ShowProgress $true

      foreach ($mailBox in $mailboxList) {
        try {
          # 5. Verificar se tem licença E3
          $license = Get-MsolUser -UserPrincipalName $mailBox | Select-Object -ExpandProperty Licenses
          if ($license.AccountSkuId -match "ENTERPRISEPACK") {
            Write-Host "[$(Get-Date)] Processando $mailBox (E3 detectada)..." -ForegroundColor Yellow

            Enable-Mailbox $mailBox -AutoExpandingArchive
            Set-Mailbox $mailBox -ElcProcessingDisabled $false
            Start-ManagedFolderAssistant $mailBox

            Write-Host "[$(Get-Date)] ✅ Concluído para $mailBox" -ForegroundColor Green
          }
          else {
            Write-Host "[$(Get-Date)] Ignorado $mailBox (sem E3)" -ForegroundColor Red
          }
        }
        catch {
          Write-Host "❌ Erro processando ${mailbox}: $($_.Exception.Message)" -ForegroundColor DarkRed
        }
      }

      # 6. Desconectar
      Disconnect-ExchangeOnline -Confirm:$false
      Write-Host "Desconectado do Exchange Online." -ForegroundColor Cyan
    }
    catch {
      Write-Host "❌ Erro geral: $($_.Exception.Message)" -ForegroundColor DarkRed
    }

    Write-Host "Aguardando 10 minutos antes da próxima execução..." -ForegroundColor Cyan
    Start-Sleep -Seconds 600
  }
}

Main
