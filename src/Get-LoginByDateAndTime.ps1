function Show-Progress {
    param (
        [Parameter(Mandatory)]
        $curr,
        [Parameter(Mandatory)]
        $total
    )

    Write-Progress -Activity "Buscando logins mais recentes..." `
        -Status "Aguarde: $curr de $total ($([Math]::Round(($curr / $total) * 100))%)" `
        -PercentComplete (($curr / $total) * 100)
}

function Get-RecentLogins {
    param (
        [Parameter(Mandatory)]
        [int]$rankingCount
    )

    $logins = Get-MgAuditLogSignIn -Top $rankingCount

    $results = @()
    $errors = @()
    $i = 0
    $loginCount = $logins.Count

    foreach ($login in $logins) {
        $i++
        Show-Progress -curr $i -total $loginCount

        try {
            $results += [PSCustomObject]@{
                UserDisplayName     = $login.UserDisplayName
                UserPrincipalName   = $login.UserPrincipalName
                CreatedDateTime     = $login.CreatedDateTime.ToLocalTime()
                Status              = $login.Status.ErrorCode -eq 0  ? "✅ Sucesso" : "❌ Erro ($($login.Status.ErrorCode))"
                ErrorCode           = $login.Status.ErrorCode
                ResourceDisplayName = $login.ResourceDisplayName
                IPAddress           = $login.IPAddress
            }
        }
        catch {
            $errors += "$($login.UserDisplayName) <$($login.UserPrincipalName)>: $($_.Exception.Message)"
        }
    }

    return $results
}

function AuditLoginWithErrors {
    param(
        [Parameter(Mandatory)]
        $login
    )

    $MSEntraIStatusCodes = @(
        [PSCustomObject]@{ ErrorCode = 16000; Title = "Interação necessária"; Description = "Conta do usuário não existe no locatário. Precisa ser adicionada como usuário externo primeiro." },
        [PSCustomObject]@{ ErrorCode = 16001; Title = "Seleção de conta inválida"; Description = "O usuário selecionou uma sessão rejeitada. Pode recuperar escolhendo outra conta." },
        [PSCustomObject]@{ ErrorCode = 16002; Title = "Seleção de sessão do aplicativo inválida"; Description = "O requisito SID especificado pelo aplicativo não foi atendido." },
        [PSCustomObject]@{ ErrorCode = 160021; Title = "Sessão solicitada pelo aplicativo não existe"; Description = "A aplicação solicitou uma sessão de usuário inexistente. Crie uma nova conta Azure para resolver." },
        [PSCustomObject]@{ ErrorCode = 16003; Title = "Conta do usuário SSO não encontrada no locatário do recurso"; Description = "Usuário não foi explicitamente adicionado ao locatário." },
        [PSCustomObject]@{ ErrorCode = 17003; Title = "Falha no provisionamento da chave de credencial"; Description = "Microsoft Entra ID não conseguiu provisionar a chave do usuário." },
        
        [PSCustomObject]@{ ErrorCode = 20001; Title = "Erro na resposta WsFed SignIn"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        [PSCustomObject]@{ ErrorCode = 20012; Title = "Mensagem WsFed inválida"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        [PSCustomObject]@{ ErrorCode = 20033; Title = "Nome do locatário do FedMetadata inválido"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        [PSCustomObject]@{ ErrorCode = 230109; Title = "Solicitações AuthN não suportadas fora do gateway"; Description = "Backup Auth Service só permite requisições AuthN via Microsoft Entra Gateway." },
        [PSCustomObject]@{ ErrorCode = 28002; Title = "Valor inválido para o escopo"; Description = "Valor fornecido para o parâmetro de escopo não é válido ao solicitar token de acesso." },
        [PSCustomObject]@{ ErrorCode = 28003; Title = "Escopo vazio"; Description = "O valor fornecido para o parâmetro de escopo não pode estar vazio ao solicitar token de acesso." },
        
        [PSCustomObject]@{ ErrorCode = 399284; Title = "Emissor do token de ID de entrada inválido"; Description = "Token de ID recebido na federação tem emissor inválido ou ausente." },
        
        [PSCustomObject]@{ ErrorCode = 40008; Title = "Erro irrecuperável no servidor do IdP OAuth2"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        [PSCustomObject]@{ ErrorCode = 40009; Title = "Erro na troca de token de atualização do IdP OAuth2"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        [PSCustomObject]@{ ErrorCode = 40010; Title = "Erro recuperável no servidor do IdP OAuth2"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        [PSCustomObject]@{ ErrorCode = 40015; Title = "Erro na troca de código de autorização do IdP OAuth2"; Description = "Problema com o provedor de identidade federado. Contate seu IDP." },
        
        [PSCustomObject]@{ ErrorCode = 50053; Title = "Conta bloqueada"; Description = "Conta temporariamente bloqueada após várias tentativas falhas" },
        [PSCustomObject]@{ ErrorCode = 50055; Title = "Senha expirada"; Description = "Usuário precisa alterar a senha" },
        [PSCustomObject]@{ ErrorCode = 50056; Title = "Nenhuma credencial fornecida"; Description = "Senha ou autenticação não foi informada" },
        [PSCustomObject]@{ ErrorCode = 50057; Title = "Conta desabilitada"; Description = "Conta de usuário desativada no Entra ID" },
        [PSCustomObject]@{ ErrorCode = 50058; Title = "Sessão inválida"; Description = "Informações de sessão não são suficientes para single-sign-on." },
        [PSCustomObject]@{ ErrorCode = 50059; Title = "Informações do locatário ausentes"; Description = "Informações do locatário não encontradas na solicitação ou nas credenciais fornecidas" },
        [PSCustomObject]@{ ErrorCode = 50072; Title = "Inscrição em MFA necessária"; Description = "Usuário precisa se inscrever para autenticação de segundo fator" },
        [PSCustomObject]@{ ErrorCode = 50074; Title = "Falha no desafio de MFA"; Description = "MFA solicitado, mas o usuário não passou" },
        [PSCustomObject]@{ ErrorCode = 50076; Title = "MFA exigido"; Description = "MFA necessário, mas não concluído" },
        [PSCustomObject]@{ ErrorCode = 50079; Title = "Atualização de segurança necessária"; Description = "Usuário precisa reenviar MFA devido a alterações de segurança" },
        [PSCustomObject]@{ ErrorCode = 50097; Title = "Autenticação de dispositivo necessária"; Description = "Autenticação de dispositivo é necessária" },
        [PSCustomObject]@{ ErrorCode = 500121; Title = "MFA não concluído"; Description = "Usuário não completou a configuração do MFA" },
        [PSCustomObject]@{ ErrorCode = 50105; Title = "Usuário não atribuído"; Description = "Usuário não tem permissão para acessar o aplicativo" },
        [PSCustomObject]@{ ErrorCode = 50125; Title = "Interrupção por redefinição de senha"; Description = "Login interrompido devido a redefinição ou registro de senha" },
        [PSCustomObject]@{ ErrorCode = 50126; Title = "Credenciais inválidas"; Description = "Senha incorreta ou usuário não existe" },
        [PSCustomObject]@{ ErrorCode = 50127; Title = "Aplicativo intermediário não instalado"; Description = "Usuário precisa instalar o aplicativo intermediário para acessar o conteúdo" },
        [PSCustomObject]@{ ErrorCode = 50129; Title = "Dispositivo não associado ao local de trabalho"; Description = "É necessário associar o dispositivo ao local de trabalho para registrar o dispositivo" },
        [PSCustomObject]@{ ErrorCode = 50132; Title = "Autenticação de risco bloqueada"; Description = "Bloqueio devido a alto risco identificado" },
        [PSCustomObject]@{ ErrorCode = 50140; Title = "Reautenticação necessária"; Description = "Sessão expirou ou precisa reautenticar" },
        [PSCustomObject]@{ ErrorCode = 50143; Title = "Sessão inválida"; Description = "Sessão inválida devido a incompatibilidade entre o locatário do usuário e a dica de domínio" },
        [PSCustomObject]@{ ErrorCode = 50144; Title = "Dispositivo não registrado"; Description = "O dispositivo do usuário não é confiável ou registrado" },
        [PSCustomObject]@{ ErrorCode = 50196; Title = "Loop detectado"; Description = "Um loop de cliente foi detectado. Verifique a lógica do aplicativo para garantir que o cache de tokens esteja implementado corretamente" },
        [PSCustomObject]@{ ErrorCode = 50199; Title = "Interrupção de segurança"; Description = "Confirmação do usuário é necessária por motivos de segurança" },
        [PSCustomObject]@{ ErrorCode = 50204; Title = "Consentimento de privacidade pendente"; Description = "Usuário externo não consentiu com a declaração de privacidade" },
        [PSCustomObject]@{ ErrorCode = 50206; Title = "Consentimento necessário"; Description = "Usuário ou administrador não consentiu em conectar ao dispositivo alvo" },
        [PSCustomObject]@{ ErrorCode = 51004; Title = "Conta de usuário não encontrada no diretório"; Description = "A conta de usuário não existe no diretório" },
        [PSCustomObject]@{ ErrorCode = 51005; Title = "Redirecionamento temporário"; Description = "Informações solicitadas estão localizadas em outro URI especificado no cabeçalho de localização" },
        [PSCustomObject]@{ ErrorCode = 51006; Title = "Reautenticação forçada"; Description = "Reautenticação forçada devido a autenticação insuficiente" },
        [PSCustomObject]@{ ErrorCode = 52004; Title = "Consentimento do LinkedIn ausente"; Description = "Usuário não forneceu consentimento para acesso aos recursos do LinkedIn" },
        [PSCustomObject]@{ ErrorCode = 53000; Title = "Bloqueado por política condicional"; Description = "Acesso bloqueado por políticas de acesso condicional (geral)" },
        [PSCustomObject]@{ ErrorCode = 53001; Title = "Política de local aplicada"; Description = "Acesso bloqueado com base na localização geográfica" },
        [PSCustomObject]@{ ErrorCode = 53002; Title = "Política de grupo aplicada"; Description = "Acesso bloqueado por não pertencer ao grupo exigido" },
        [PSCustomObject]@{ ErrorCode = 53003; Title = "Bloqueado por política condicional"; Description = "Acesso bloqueado por regras de localização, grupo, dispositivo, etc" },
        [PSCustomObject]@{ ErrorCode = 53004; Title = "Configuração de MFA bloqueada"; Description = "Não é possível configurar métodos de MFA devido a atividade suspeita" },
        [PSCustomObject]@{ ErrorCode = 53005; Title = "Políticas de proteção do Intune necessárias"; Description = "Aplicativo precisa impor políticas de proteção do Intune" },
        [PSCustomObject]@{ ErrorCode = 53006; Title = "Autenticação federada necessária"; Description = "Autenticação requerida de um provedor de identidade federado" },
        [PSCustomObject]@{ ErrorCode = 53008; Title = "Navegador não suportado"; Description = "O navegador utilizado não é suportado para autenticação" },
        [PSCustomObject]@{ ErrorCode = 53010; Title = "Configuração de MFA restrita"; Description = "Organização exige que a configuração de MFA seja feita de locais ou dispositivos específicos" },
        [PSCustomObject]@{ ErrorCode = 54000; Title = "Restrição de idade legal"; Description = "Usuário não tem idade legal para acessar o aplicativo" },
        [PSCustomObject]@{ ErrorCode = 54005; Title = "Código de autorização já utilizado"; Description = "Código de autorização OAuth2 já foi resgatado" },
        [PSCustomObject]@{ ErrorCode = 54008; Title = "MFA necessário com credencial não suportada"; Description = "MFA é necessário, mas a credencial usada não é suportada como primeiro fator" },
        
        [PSCustomObject]@{ ErrorCode = 65001; Title = "Consentimento do aplicativo ausente"; Description = "Usuário ou administrador não consentiu em usar o aplicativo com o ID especificado" },
        [PSCustomObject]@{ ErrorCode = 65001; Title = "Delegação não existe"; Description = "O usuário ou administrador não consentiu o uso da aplicação $($login.ResourceDisplayName). Envie uma requisição interativa de autorização para esse usuário e recurso." },
        [PSCustomObject]@{ ErrorCode = 65002; Title = "Consentimento entre app e recurso deve ser configurado via pré-autorização"; Description = "Aplicações da Microsoft precisam de aprovação do proprietário da API antes de solicitar tokens para essa API. Um desenvolvedor pode estar tentando reutilizar um App ID da Microsoft, o que é impedido para evitar falsificação." },
        [PSCustomObject]@{ ErrorCode = 65004; Title = "Usuário recusou o consentimento"; Description = "O usuário recusou o consentimento para acessar o app. Solicite que ele tente novamente e conceda o consentimento." },
        [PSCustomObject]@{ ErrorCode = 65005; Title = "Aplicação mal configurada"; Description = "A lista de acesso a recursos do app não contém apps descobertos pelo recurso, ou o app solicitou acesso a recurso não especificado. Pode ser configuração incorreta no identificador (Entity) se usar SAML." },
        [PSCustomObject]@{ ErrorCode = 650052; Title = "App precisa acessar serviço não habilitado"; Description = "O app precisa acessar um serviço que sua organização não assinou ou habilitou. Contate o administrador de TI para revisar a configuração das assinaturas de serviço." },
        [PSCustomObject]@{ ErrorCode = 650054; Title = "Recurso solicitado não está disponível"; Description = "A aplicação solicitou permissões para um recurso removido ou indisponível. Verifique se todos os recursos que o app chama existem no locatário onde está operando." },
        [PSCustomObject]@{ ErrorCode = 650056; Title = "Aplicação mal configurada"; Description = "Pode ser porque o cliente não listou permissões para '{name}', ou o administrador não consentiu no locatário, ou identificador do app está errado, ou certificado inválido. Contate o admin para corrigir a configuração ou consentir pelo locatário." },
        [PSCustomObject]@{ ErrorCode = 650057; Title = "Recurso inválido"; Description = "O cliente pediu acesso a um recurso que não está listado nas permissões requisitadas no registro do app. IDs e nomes dos apps e recursos são listados para referência." },
        [PSCustomObject]@{ ErrorCode = 67003; Title = "Ator não é identidade de serviço válida"; Description = "A identidade solicitante não é uma identidade de serviço válida." }
        
        [PSCustomObject]@{ ErrorCode = 700016; Title = "Cliente não suportado"; Description = "Tentativa de login via cliente ou app bloqueado pelas políticas" },
        [PSCustomObject]@{ ErrorCode = 7000218; Title = "Senha incorreta"; Description = "Senha digitada incorretamente durante o login" },
        [PSCustomObject]@{ ErrorCode = 7000228; Title = "Login externo bloqueado"; Description = "Tentativa de login de fora da organização foi bloqueada" },
        [PSCustomObject]@{ ErrorCode = 70044; Title = "Conta bloqueada por política de identidade"; Description = "Bloqueio por risco, senha comprometida ou política condicional" },
        [PSCustomObject]@{ ErrorCode = 70046; Title = "Sessão expirada"; Description = "Sessão expirada ou verificação de reautenticação falhou" },
        [PSCustomObject]@{ ErrorCode = 70049; Title = "Dispositivo não em conformidade"; Description = "Dispositivo fora das regras de conformidade do Intune" },
        [PSCustomObject]@{ ErrorCode = 70000; Title = "Concessão Inválida"; Description = "Falha na autenticação. O token de atualização (refresh token) não é válido. Pode ocorrer por: cabeçalho de vinculação do token vazio ou hash da vinculação do token não corresponde." },
        [PSCustomObject]@{ ErrorCode = 70001; Title = "Cliente Não Autorizado"; Description = "A aplicação está desabilitada. Para mais informações, veja o artigo de resolução do erro AADSTS70001." },
        [PSCustomObject]@{ ErrorCode = 700011; Title = "Aplicativo Cliente Não Encontrado no Locatário OrgID"; Description = "Aplicação com identificador {appIdentifier} não encontrada no diretório. Um app cliente solicitou um token do seu locatário, mas o app não existe no locatário, causando falha na chamada." },
        [PSCustomObject]@{ ErrorCode = 70002; Title = "Cliente Inválido"; Description = "Erro ao validar as credenciais. O client_secret especificado não corresponde ao esperado para esse cliente. Corrija o client_secret e tente novamente." },
        [PSCustomObject]@{ ErrorCode = 700025; Title = "Cliente Público Inválido com Credencial"; Description = "O cliente é público, portanto nem 'client_assertion' nem 'client_secret' devem ser apresentados." },
        [PSCustomObject]@{ ErrorCode = 700027; Title = "Falha na Validação da Assinatura da Declaração do Cliente"; Description = "Erro do desenvolvedor - o app tenta se autenticar sem os parâmetros corretos ou necessários." },
        [PSCustomObject]@{ ErrorCode = 70003; Title = "Tipo de Concessão Não Suportado"; Description = "O app retornou um tipo de concessão não suportado." },
        [PSCustomObject]@{ ErrorCode = 700030; Title = "Certificado Inválido"; Description = "O nome do assunto no certificado não está autorizado. Nomes de assunto/nomes alternativos (até 10) no certificado do token são: {certificateSubjects}." },
        [PSCustomObject]@{ ErrorCode = 70004; Title = "URI de Redirecionamento Inválido"; Description = "O app retornou um URI de redirecionamento inválido. O endereço especificado não corresponde a nenhum configurado ou na lista aprovada do OIDC." },
        [PSCustomObject]@{ ErrorCode = 70005; Title = "Tipo de Resposta Não Suportado"; Description = "O app retornou um tipo de resposta não suportado devido a: tipo 'token' não habilitado para o app, ou tipo 'id_token' requer escopo 'OpenID' ou contém parâmetro OAuth inválido no wctx codificado." },
        [PSCustomObject]@{ ErrorCode = 700054; Title = "Response_type 'id_token' não habilitado"; Description = "O app solicitou um token ID do endpoint de autorização, mas não tem concessão implícita de token ID habilitada. Acesse o portal Entra para habilitar." },
        [PSCustomObject]@{ ErrorCode = 70007; Title = "Modo de Resposta Não Suportado"; Description = "O app retornou um valor não suportado de response_mode ao solicitar um token." },
        [PSCustomObject]@{ ErrorCode = 70008; Title = "Concessão Expirada ou Revogada"; Description = "O token de atualização expirou devido à inatividade. O token foi emitido em XXX e esteve inativo por determinado tempo." },
        [PSCustomObject]@{ ErrorCode = 700082; Title = "Concessão Expirada ou Revogada por Inatividade"; Description = "O token de atualização expirou por inatividade. O token foi emitido em {issueDate} e ficou inativo por {time}. É esperado que o usuário não use o app por muito tempo e o token expire quando o app tentar renovar." },
        [PSCustomObject]@{ ErrorCode = 700084; Title = "Token de atualização para SPA expirado"; Description = "O token de atualização emitido para um SPA tem vida limitada e fixa de {time}, que não pode ser estendida. Está expirado e é necessário novo login. Emitido em {issueDate}." },
        [PSCustomObject]@{ ErrorCode = 70011; Title = "Escopo Inválido"; Description = "O escopo solicitado pelo app é inválido." },
        [PSCustomObject]@{ ErrorCode = 70012; Title = "Erro no Servidor MSA"; Description = "Erro no servidor ao autenticar um usuário MSA (consumer). Tente novamente ou abra um chamado de suporte." },
        [PSCustomObject]@{ ErrorCode = 70016; Title = "Autorização Pendente"; Description = "Erro no fluxo OAuth 2.0 device flow. A autorização está pendente. O dispositivo tentará novamente." },
        [PSCustomObject]@{ ErrorCode = 70018; Title = "Código de Verificação Inválido"; Description = "Código inválido por erro do usuário ao digitar o código para o device flow. A autorização não foi aprovada." },
        [PSCustomObject]@{ ErrorCode = 70019; Title = "Código Expirado"; Description = "Código de verificação expirado. Solicite que o usuário tente novamente o login." },
        [PSCustomObject]@{ ErrorCode = 70043; Title = "Token Inválido devido à frequência de login"; Description = "O token de atualização expirou ou é inválido devido às verificações de frequência de login por Conditional Access. Emitido em {issueDate}, tempo máximo permitido é {time}." },
        [PSCustomObject]@{ ErrorCode = 75001; Title = "Erro de Serialização de Binding"; Description = "Erro ocorreu durante o binding da mensagem SAML." },
        [PSCustomObject]@{ ErrorCode = 75003; Title = "Erro de Binding Não Suportado"; Description = "O app retornou erro relacionado a binding não suportado (resposta SAML não pode ser enviada via bindings que não HTTP POST)." },
        [PSCustomObject]@{ ErrorCode = 75005; Title = "Mensagem SAML Inválida"; Description = "Microsoft Entra não suporta o pedido SAML enviado pelo app para SSO. Veja o artigo de resolução para erro AADSTS75005." },
        [PSCustomObject]@{ ErrorCode = 7500514; Title = "Tipo de resposta SAML suportado não encontrado"; Description = "Tipos suportados são 'Response' (XML namespace urn:oasis:names:tc:SAML:2.0:protocol) ou 'Assertion' (XML namespace urn:oasis:names:tc:SAML:2.0:assertion). Erro de aplicação a ser tratado pelo desenvolvedor." },
        [PSCustomObject]@{ ErrorCode = 750054; Title = "Parâmetros SAMLRequest ou SAMLResponse ausentes"; Description = "SAMLRequest ou SAMLResponse devem estar presentes como parâmetros na query string para binding SAML Redirect. Veja artigo de resolução para erro AADSTS750054." },
        [PSCustomObject]@{ ErrorCode = 75008; Title = "Erro de Requisição Negada"; Description = "Pedido do app foi negado porque a requisição SAML tinha destino inesperado." },
        [PSCustomObject]@{ ErrorCode = 75011; Title = "Contexto de Autenticação Não Correspondente"; Description = "Método de autenticação do usuário não corresponde ao método solicitado. Veja artigo de resolução para erro AADSTS75011." },
        [PSCustomObject]@{ ErrorCode = 75016; Title = "NameIDPolicy Inválido na Requisição de Autenticação SAML2"; Description = "A requisição de autenticação SAML2 tem NameIDPolicy inválido." },
        [PSCustomObject]@{ ErrorCode = 76021; Title = "Aplicação Requer Requisições Assinadas"; Description = "Requisição enviada pelo cliente não está assinada, enquanto a aplicação exige requisições assinadas." },
        [PSCustomObject]@{ ErrorCode = 76026; Title = "Tempo de Emissão da Requisição Expirado"; Description = "IssueTime na requisição de autenticação SAML2 está expirado." }
        
        [PSCustomObject]@{ ErrorCode = 81010; Title = "Conta bloqueada pelo Identity Protection"; Description = "Login suspeito ou risco identificado pelo Microsoft Entra" },
        [PSCustomObject]@{ ErrorCode = 81012; Title = "Risco de login"; Description = "Login considerado arriscado e bloqueado pelo Identity Protection" },

        [PSCustomObject]@{ ErrorCode = 80001; Title = "Armazenamento Local Indisponível"; Description = "O Agente de Autenticação não consegue se conectar ao Active Directory. Verifique se os servidores estão no mesmo domínio que os usuários e se há conectividade." }
        [PSCustomObject]@{ ErrorCode = 80002; Title = "Tempo Esgotado na Validação de Senha"; Description = "A solicitação de validação de senha expirou. Verifique se o Active Directory está acessível e respondendo aos agentes." }
        [PSCustomObject]@{ ErrorCode = 80005; Title = "Erro Web Imprevisível"; Description = "Erro desconhecido ao processar a resposta do Agente de Autenticação. Tente novamente. Se persistir, abra um chamado de suporte." }
        [PSCustomObject]@{ ErrorCode = 80007; Title = "Erro na Validação de Senha Local"; Description = "O Agente de Autenticação não conseguiu validar a senha do usuário. Verifique os logs do agente e o funcionamento do Active Directory." }
        [PSCustomObject]@{ ErrorCode = 80010; Title = "Erro de Criptografia na Validação"; Description = "O Agente de Autenticação não conseguiu descriptografar a senha." }
        [PSCustomObject]@{ ErrorCode = 80012; Title = "Login Fora do Horário Permitido"; Description = "O usuário tentou fazer login fora do horário permitido conforme definido no Active Directory." }
        [PSCustomObject]@{ ErrorCode = 80013; Title = "Diferença de Horário (Time Skew)"; Description = "Falha na autenticação devido à diferença de horário entre o agente de autenticação e o Active Directory. Corrija o problema de sincronização de tempo." }
        [PSCustomObject]@{ ErrorCode = 80014; Title = "Tempo Excedido pelo Agente de Autenticação"; Description = "A resposta de validação excedeu o tempo máximo permitido. Abra um chamado com o código de erro, ID de correlação e timestamp." }
        [PSCustomObject]@{ ErrorCode = 81004; Title = "Falha na Autenticação Kerberos"; Description = "A tentativa de autenticação Kerberos falhou." }
        [PSCustomObject]@{ ErrorCode = 81005; Title = "Pacote de Autenticação Não Suportado"; Description = "O pacote de autenticação não é suportado." }
        [PSCustomObject]@{ ErrorCode = 81006; Title = "Cabeçalho de Autorização Ausente"; Description = "Nenhum cabeçalho de autorização foi encontrado." }
        [PSCustomObject]@{ ErrorCode = 81007; Title = "Tenant Não Habilitado para SSO"; Description = "O tenant não está habilitado para SSO transparente (Seamless SSO)." }
        [PSCustomObject]@{ ErrorCode = 81009; Title = "Formato Inválido no Cabeçalho de Autorização"; Description = "Não foi possível validar o ticket Kerberos do usuário." }
        [PSCustomObject]@{ ErrorCode = 81010; Title = "Token Kerberos Inválido ou Expirado"; Description = "Falha no SSO transparente porque o ticket Kerberos do usuário está expirado ou inválido." }
        [PSCustomObject]@{ ErrorCode = 81011; Title = "Usuário Não Encontrado pelo SID"; Description = "Não foi possível localizar o objeto do usuário com base nas informações do ticket Kerberos." }
        [PSCustomObject]@{ ErrorCode = 81012; Title = "UPN do Token Difere do UPN Escolhido"; Description = "O usuário tentando fazer login é diferente do usuário conectado no dispositivo." }
        
        [PSCustomObject]@{ ErrorCode = 90025; Title = "Erro interno de serviço"; Description = "Serviço interno da Microsoft Entra atingiu seu limite de tentativas de login" }
    )

    $matched = $MSEntraIStatusCodes | Where-Object { $_.ErrorCode -eq $login.ErrorCode }

    if (!$matched) {
        Write-Host "❌ $($login.UserDisplayName) <$($login.UserPrincipalName)> tentou acessar '$($login.ResourceDisplayName)' em $($login.CreatedDateTime) a partir do IP $($login.IPAddress) — Resultado: $($login.Status) ($($login.StatusCode) - Código desconhecido)"
    }
    
    Write-Host "❌ $($login.UserDisplayName) <$($login.UserPrincipalName)> tentou acessar '$($login.ResourceDisplayName)' em $($login.CreatedDateTime) a partir do IP $($login.IPAddress) — Resultado: $($login.Status) ($($login.StatusCode) - $($matched.Title): $($matched.Description))"
}

function Main {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║              🪟 Microsoft Entra ID - Últimos Logins                ║" -ForegroundColor Cyan
    Write-Host "║--------------------------------------------------------------      ║" -ForegroundColor Cyan
    Write-Host "║ Autor      : Natan Felipe de Oliveira                              ║" -ForegroundColor Cyan
    Write-Host "║ Descrição  : Mostra os últimos logins indicando o serviço acessado ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "🔄 Conectando ao Microsoft Entra ID..."
    Connect-MgGraph -Scopes "AuditLog.Read.All"  -NoWelcome
    Write-Host "✅ Conectado com sucesso!`n"

    [int]$rankingCount = Read-Host "Quantos resultados deseja visualizar no ranking"

    if (-not $rankingCount -or $rankingCount -lt 1) {
        Write-Host "❌ Insira um valor maior que 0!" -ForegroundColor Red
        exit 1
    }

    $loginResults = Get-RecentLogins -rankingCount $rankingCount
    
    $rankedOutput = $loginResults | Select-Object -First $rankingCount | ForEach-Object -Begin { $i = 1 } -Process {
        [PSCustomObject]@{
            Rank           = $i
            "Nome"         = $_.UserDisplayName
            "UPN"          = $_.UserPrincipalName
            "Último login" = $_.CreatedDateTime
            "Status"       = $_.Status
            "App usado"    = $_.ResourceDisplayName
            "IP"           = $_.IPAddress
        }
        $i++
    }

    $rankedOutput | Format-Table -AutoSize
    
    foreach ($result in $loginResults) {
        if ($result.ErrorCode -ne 0) {
            AuditLoginWithErrors -login $result
        }
    }
}

try {
    Main
}
catch {
    Write-Host "❌ Ocorreu um erro inesperado: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
