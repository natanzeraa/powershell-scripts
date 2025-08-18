# 📄 Habilitar Arquivo Morto Expansivo + MRM no Exchange Online

## 📝 Descrição Geral

Este script automatiza a **habilitação do Arquivo Morto Expansivo (Auto-Expanding Archive)** e a **ativação do MRM (Messaging Records Management)** no **Exchange Online**, garantindo a aplicação de políticas de retenção e otimização do uso de espaço em caixas de correio corporativas.

Ele foi desenvolvido em PowerShell, compatível com versões **5.1 e 7+**, e executa de forma **contínua em loop** com intervalos de **10 minutos**.

---

## ⚙️ Funcionalidades Principais

1. **Ajuste temporário da política de execução**

   * Define a execução do script como `RemoteSigned` apenas para a sessão atual.
   * Evita falhas de execução por restrições de política do PowerShell.

2. **Verificação e instalação de módulos obrigatórios**

   * **ExchangeOnlineManagement**
   * **MSOnline**
   * Caso não estejam instalados, são baixados e importados automaticamente.

3. **Conexão com Exchange Online**

   * Usa a conta administrativa definida em `$UserUPN`.
   * Conecta ao serviço MSOL (`Connect-MsolService`).
   * Conecta ao Exchange Online (`Connect-ExchangeOnline`).

4. **Processamento de mailboxes**

   * Itera sobre a lista de caixas de correio definida em `$mailboxList`.
   * Para cada usuário:

     * **Verifica se possui licença E3 (ENTERPRISEPACK)**.
     * Se **SIM**:

       * Habilita o **Arquivo Morto Expansivo** (`Enable-Mailbox -AutoExpandingArchive`).
       * Reativa o **Managed Folder Assistant** para processamento de políticas (`Set-Mailbox -ElcProcessingDisabled $false` + `Start-ManagedFolderAssistant`).
       * Exibe log de sucesso com data/hora.
     * Se **NÃO**:

       * Ignora a caixa de correio, exibindo log em vermelho.

5. **Tratamento de erros**

   * Usa `try/catch` para capturar e exibir erros detalhados de cada mailbox.
   * Caso haja falha geral, o erro é registrado sem interromper a execução contínua.

6. **Desconexão e ciclo contínuo**

   * Após processar todas as caixas, desconecta do Exchange Online (`Disconnect-ExchangeOnline`).
   * Aguardar **10 minutos** (`Start-Sleep -Seconds 600`).
   * Reinicia o processo automaticamente (loop infinito).

---

## 🔐 Pré-requisitos

* PowerShell 5.1 ou 7+.
* Permissões administrativas no **Exchange Online**.
* Conta administrativa configurada em `$UserUPN`.
* Usuários precisam possuir licença **Office 365 E3 (ENTERPRISEPACK)** para que o Arquivo Morto Expansivo esteja disponível.

---

## 📊 Benefícios e Objetivo

* Garante que caixas de correio **não ultrapassem os limites de espaço**.
* Automatiza a **migração de itens antigos para o Arquivo Morto Expansivo**.
* Assegura que o **MRM (políticas de retenção)** esteja ativo e processando corretamente.
* Evita trabalho manual de habilitação individual para dezenas de usuários.
* Mantém execução contínua, garantindo que novos usuários adicionados à lista sejam processados automaticamente.

---

## 🚀 Fluxo Resumido

1. Ajusta política de execução.
2. Garante módulos instalados.
3. Conecta ao Exchange Online.
4. Percorre lista de caixas de correio:

   * Verifica licença E3.
   * Habilita Arquivo Morto Expansivo e ativa MRM.
5. Desconecta.
6. Aguarda 10 minutos.
7. Repete o ciclo automaticamente.

---

### 🔎 Como o script funciona

1. **Cabeçalho / Inicialização**

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
   ```

   👉 Ajusta a política de execução apenas para a sessão atual. Isso evita bloqueios no script.

---

2. **Verificação e instalação dos módulos**

   ```powershell
   if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
     Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
   }
   if (-not (Get-Module -ListAvailable -Name MSOnline)) {
     Install-Module -Name MSOnline -Scope CurrentUser -Force
   }

   Import-Module ExchangeOnlineManagement
   Import-Module MSOnline
   ```

   👉 Garante que os módulos necessários estão disponíveis e os importa.

---

3. **Lista de caixas de correio**

   ```powershell
   $mailboxList = @("email11@...", "email27@...", ...)
   $UserUPN = "user@email.com.br"
   ```

   👉 Define **quais mailboxes** serão processadas e **qual conta administrativa** será usada para autenticar.

---

4. **Loop contínuo**

   ```powershell
   while ($true) {
     ...
     Start-Sleep -Seconds 600
   }
   ```

   👉 O script **nunca para**. Ele roda, processa, espera **10 minutos** e roda de novo.

---

5. **Conexão aos serviços**

   ```powershell
   Connect-MsolService
   Connect-ExchangeOnline -UserPrincipalName $UserUPN -ShowProgress $true
   ```

   👉 Conecta no **MSOnline** e no **Exchange Online** usando a conta definida.

---

6. **Processamento por mailbox**

   ```powershell
   $license = Get-MsolUser -UserPrincipalName $mailBox | Select-Object -ExpandProperty Licenses
   if ($license.AccountSkuId -match "ENTERPRISEPACK") {
     Enable-Mailbox $mailBox -AutoExpandingArchive
     Set-Mailbox $mailBox -ElcProcessingDisabled $false
     Start-ManagedFolderAssistant $mailBox
   }
   ```

   👉 Para cada usuário da lista:

   * Verifica se tem licença **E3 (ENTERPRISEPACK)**.
   * Se sim:

     * **Habilita Arquivo Morto Expansivo**.
     * **Ativa o processamento de políticas de retenção (MRM)**.
     * **Força execução imediata do Managed Folder Assistant**.
   * Se não: ignora e mostra log.

---

7. **Desconexão**

   ```powershell
   Disconnect-ExchangeOnline -Confirm:$false
   ```

   👉 Desconecta a sessão do Exchange antes de iniciar o próximo ciclo.

---

### ✅ Conclusão

Ou seja, o que o **script faz** é:

* **Verificar módulos e instalar se necessário**.
* **Conectar ao Exchange Online com conta administrativa**.
* **Percorrer a lista de mailboxes** definida manualmente.
* Para cada mailbox com licença **E3**:

  * Habilitar **Arquivo Morto Expansivo (AutoExpandingArchive)**.
  * Ativar o **Managed Folder Assistant** (para aplicar regras de MRM).
* **Rodar continuamente a cada 10 minutos**, repetindo o processo.
