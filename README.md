# NotasSimples Mobile

A única versão publicada do aplicativo mobile está em `/restrito/`. A página principal encaminha para ela e não oferece login nem acesso ao OneDrive.

A autenticação da conta Microsoft pessoal solicita somente `Files.ReadWrite.AppFolder`. O aplicativo usa `Dados/Historico` e `Dados/Textos` na pasta `Apps/Notas Simples Restrito` do OneDrive; não cria a pasta `Dados` automaticamente. O redirect SPA cadastrado no Microsoft Entra ID continua em `https://raphaelrezenderadio.github.io/notas-simples/restrito/`.

O programa Windows 2.2 e seus dados não são publicados neste repositório. O modo de teste da antiga versão mobile foi retirado. Para evitar perda de rascunhos que ainda não tenham sido sincronizados, chaves antigas de armazenamento local são preservadas, mas não são usadas pelo app final. A página principal elimina as credenciais de sessão do aplicativo antigo quando é aberta.
