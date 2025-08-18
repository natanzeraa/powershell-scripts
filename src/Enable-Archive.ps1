param(
  [Parameter(Mandatory)] $adminUPN
)

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
  if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
    Write-Host "Instalando módulo Microsoft.Graph..." -ForegroundColor Yellow
  }

  # Import-Module Microsoft.Graph
  # Import-Module ExchangeOnlineManagement

  if (-not (Get-Module -Name ExchangeOnlineManagement)) {
    Import-Module ExchangeOnlineManagement
  }

  # if (-not (Get-Module -Name Microsoft.Graph)) {
  #   Import-Module Microsoft.Graph
  # }

  # 3. Lista de emails a serem percorridos
  $mailboxList = Get-Content -Path ".\temp\mail_list.txt"

  while ($true) {
    try {
      # 4. Conectar ao Exchange Online
      Connect-ExchangeOnline -UserPrincipalName $adminUPN -ShowProgress $true
      Write-Host "Conectado ao Exchange Online como $adminUPN" -ForegroundColor Yellow

      # 5. Conectar ao Microsoft Graph
      Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
      Write-Host "Conectado ao Microsoft Graph para verificação de licenças." -ForegroundColor Yellow

      # 6. Buscar SKUs disponíveis (para encontrar o GUID da E3)
      $skus = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId
      $e3Sku = ($skus | Where-Object { $_.SkuPartNumber -eq "ENTERPRISEPACK" }).SkuId

      foreach ($mailBox in $mailboxList) {
        try {
          # 7. Verificar se o usuário tem licença E3
          # $user = Get-MgUser -UserId $mailBox -Property "AssignedLicenses"

          # $mb = Get-Mailbox $mailBox -ErrorAction Stop
          # $user = Get-MgUser -UserId $mb.ExternalDirectoryObjectId -Property "assignedLicenses"

          $user = Get-MgUser -Filter "mail eq '$mailBox'" -Property "assignedLicenses"
          $hasE3 = $user.AssignedLicenses.SkuId -contains $e3Sku

          if ($hasE3) {
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
          Write-Host "❌ Erro processando ${mailBox}: $($_.Exception.Message)" -ForegroundColor DarkRedllll
        }
      }

      # 8. Desconectar
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
