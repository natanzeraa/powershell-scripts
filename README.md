<img src="https://skillicons.dev/icons?i=powershell&perline=10" alt="Sistemas Operacionais" />

## Powershell Scripts

> Este repositório tem como objetivo armazenar scripts Powershell para gerenciamento e auditoria de dados do Microsoft Entra ID
---
### Como executar os scripts
Antes executar qualquer um desses scripts certifique-se de ter instaldo os modulos do MG Graph

```	
Get-InstalledModule | Where-Object {$_.Name -match "Microsoft.Graph"}
```

Se não retornar nada na saída do seu terminal, execute o comando abaixo:

```	
Install-Module -Name "Microsoft.Graph"
```
**OBS: _Isso fará com que os módulos sejam instalados globalmente._**

Se prefirir instalar apenas para o usuário atual

```	
Install-Module Microsoft.Graph -Scope CurrentUser
```

Se quiser saber mais informações sobre como atualizar os módulos acesse:

- [**How to connect to microsoft graph api from Powershell 📍**](https://www.sharepointdiary.com/2023/04/how-to-connect-to-microsoft-graph-api-from-powershell.html)

- [**Microsoft graph authentication 📍**](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph?view=graph-powershell-1.0)