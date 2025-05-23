# 📦 Coleta e Ranking de Caixas de E-mail no Exchange Online

[Acesse o script clicando aqui 📍](https://github.com/natanzeraa/scripts-and-automation/blob/main/PowerShell%20Scripts/MicrosoftEntraID/GetExchangeMailBoxSize.ps1)

Este script PowerShell coleta estatísticas de caixas de e-mail no Exchange Online e gera um ranking das caixas mais ocupadas, exportando os resultados em CSV.

### 📋 Funcionalidades

* Lista todas as caixas de e-mail da organização.
* Exibe o nome da organização e total de caixas.
* Permite ao usuário escolher quantas caixas exibir no ranking.
* Exibe uma barra de progresso durante a execução.
* Gera relatório das caixas de email com as seguintes informações:
  * Nome do usuário
  * E-mail
  * Qtd. em uso
  * Capacidade total
  * Porcentagem de uso
  * Espaço disponível
  * Quantidade de e-mails enviados
* Exporta os dados em CSV para a pasta `output`.

### ✅ Pré-requisitos

> Para que todas as funcionalidades sejam executadas e funcionem da maneira correta, execute o script em um ambiente **PowerShell 7.5.1**

O ambiente também precisa atender aos seguintes requisitos:

#### 1. **Permissões e Conectividade**

* A conta que executa o script deve ter permissões administrativas no Exchange Online.
* O PowerShell deve estar conectado ao Exchange Online (via módulo do Exchange Online Management).

```powershell
Connect-ExchangeOnline -UserPrincipalName "seu_usuario@dominio.com" -ShowBanner:$false
```

#### 2. **Módulo PowerShell Necessário**

Instale o módulo **ExchangeOnlineManagement** caso ainda não o tenha:

```powershell
Install-Module ExchangeOnlineManagement
```

#### 3. **Permissão para Execução de Scripts**

Certifique-se de que a execução de scripts está habilitada:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### ▶️ Como usar

1. Abra o PowerShell como Administrador.
2. Conecte-se ao Exchange Online:

```powershell
Connect-ExchangeOnline -UserPrincipalName seu_usuario@dominio.com
```

3. Execute o script:

```powershell
.\GetExchangeMailBoxSize.ps1
```

4. Quando solicitado, informe o número de caixas de e-mail mais ocupadas que deseja ver no ranking (por exemplo, `10`).

### 📁 Exportação

* O relatório será salvo automaticamente na pasta `output`, localizada um nível acima do diretório do script.
* Nome do arquivo gerado: `top_X_caixas_de_email.csv`, onde `X` é o número informado pelo usuário.

### ⚠️ Possíveis erros tratados

* Caixas de e-mail sem dados de tamanho.
* Falha ao coletar estatísticas individuais (exibido ao final da execução).

### 📌 Observações

* O script é útil para administradores de ambientes Microsoft 365 que desejam monitorar o uso de caixas de e-mail.
* Pode ser adaptado para rodar via agendamento ou integração com outras ferramentas de auditoria.

---

### 🔍 Como o script funciona
Aqui está todo o código do seu script dentro de um único bloco, utilizando a estrutura `<details>` para cada parte do processo. Isso vai criar uma documentação organizada e interativa, onde cada etapa pode ser expandida e contraída conforme necessário.

<details>
  <summary>📥 Iniciando contagem de caixas de e-mail</summary>

  Aqui (1° linha) buscamos todas as caixas disponíveis da organização (que não são compartilhadas e são do tipo "Member", todos os convidados (Guests) ficam de fora). Depois (2° linha) somamos todas as caixas disponíveis para usar durante a execução do script.

  ```powershell
  $mailboxes = Get-Mailbox -ResultSize Unlimited
  $mailboxesCount = $mailboxes.Count
```
</details>

<details>
  <summary>🏢 Exibindo informações da organização</summary>

Exibe o nome da organização e a quantidade de caixas coletadas.

```powershell
$orgName = (Get-OrganizationConfig).DisplayName
Write-Host "`nOrganização: $orgName"
Write-Host "`nTotal de caixas de e-mail: $mailboxesCount"
```
</details>

<details>
  <summary>🔢 Escolha do ranking</summary>

Permite ao usuário informar quantas caixas ele quer ver no ranking.

```powershell
[int]$topRankingCount = Read-Host "`nQuantas caixas de e-mail mais ocupadas você deseja visualizar no ranking"
```

</details>

<details>
  <summary>📊 Função de progresso</summary>

Exibe a barra de progresso durante a coleta das caixas de e-mail.

```powershell
function Show-Progress($current, $total) {
    Write-Progress -Activity "Coletando caixas de e-mail" `
        -Status "$current de $total processado(s) ($([math]::Round(($current / $total) * 100))%)" `
        -PercentComplete (($current / $total) * 100)
}
```

</details>

<details>
  <summary>📏 Funções de medição</summary>

Calcula o total de uso, o ranking das caixas mais ocupadas e a média do uso por caixa.

```powershell
function Measure-AllMailboxesSize { ... }
function Measure-RankedMailboxesSize { ... }
function Measure-MailboxesSizeMean { ... }
```

</details>

<details>
  <summary>🔄 Conversão de cotas</summary>

Converte o valor de `ProhibitSendQuota` (o tamanho máximo da caixa) para bytes para realizar os cálculos corretamente.

```powershell
function Convert-StringToBytes { ... }
```

</details>

<details>
  <summary>📈 Coleta e ranking das caixas</summary>

Coleta as estatísticas das caixas com `Get-MailboxStatistics` e gera o ranking das mais ocupadas.

```powershell
function Get-MailboxUsageReport { ... }
```

</details>

<details>
  <summary>📁 Exportação para CSV</summary>

Cria a pasta `output` e exporta o ranking das caixas mais ocupadas em um arquivo CSV.

```powershell
$csvDir = Join-Path $PSScriptRoot "..\output"
$csvPath = Join-Path $csvDir "top_${topRankingCount}_caixas_de_email.csv"
Export-Csv -Path $csvPath ...
```

</details>

<details>
  <summary>⏱️ Duração e erros</summary>

Exibe o tempo de execução do script e erros encontrados, se houver.

```powershell
$duration = (Get-Date) - $startTime
Write-Host "`nDuração: ..."
if ($errors.Count -gt 0) { ... }
```
</details>
