# Despesas Ana e Thiago

App de controle de despesas do casal. Le e grava a planilha diretamente no OneDrive
pessoal via Microsoft Graph, com login da conta Microsoft.

**Nenhum dado financeiro fica neste repositorio.** O que esta aqui e apenas a interface.
A planilha continua privada no OneDrive e so e acessivel por quem tem permissao nela.

Endereco: https://thiagoindahouse.github.io/despesas-app/

## Como funciona

- Login OAuth2 com PKCE, feito sem biblioteca externa
- Leitura e escrita do .xlsx implementadas na mao, com CompressionStream nativo
- Controle de concorrencia por eTag: se o outro editar no meio, o app recusa gravar
- Funciona tambem com arquivo local, como alternativa
