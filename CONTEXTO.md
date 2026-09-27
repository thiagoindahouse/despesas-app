# Contexto do projeto

Registro do que este app é, por que cada decisão foi tomada e onde estão as
armadilhas. Escrito para quem pegar isto daqui a seis meses — inclusive eu.

---

## O que é

App web de página única para Thiago e Ana controlarem as despesas da casa.
Não tem servidor, não tem banco, não tem dependência externa. É um HTML só que
abre a planilha do OneDrive, lê, mostra e grava de volta.

- **App:** https://thiagoindahouse.github.io/despesas-app/
- **Planilha:** OneDrive pessoal de `thiagopaiva@live.com`
  - drive `A2C908BE3C24ADBC`
  - pasta `/Despesas Ana e Thiago/`
  - arquivo `Despesas Ana e Thiago.xlsx`
- **Acesso:** Thiago é dono; Ana (`ana_santoliquido@hotmail.com`) tem `write`.
  Não existe link público — havia um de edição anônima e foi removido em 27/09/2026.

---

## A regra do acerto

Esta é a regra do negócio. Tudo no app existe para calcular isto certo.

Cada despesa é dividida **meio a meio**. Quem paga o boleto fica **credor da
metade**:

| Situação | Resultado |
|---|---|
| **Thiago** paga R$ 1.000 | **Ana deve R$ 500** a ele |
| **Ana** paga R$ 1.000 | **Thiago deve R$ 500** a ela |
| **Casa** paga R$ 1.000 | ninguém deve nada |

No fim do mês soma tudo: quem pagou mais **recebe** a diferença.

Existem exceções onde a despesa é 100% de um dos dois — daí os modos
"Conta do Thiago" e "Conta da Ana".

### O bug que originou a reescrita desta parte

A primeira versão tinha uma coluna chamada **`% Thiago`**, que significava
*"quanto da despesa cabe ao Thiago"*. O Thiago leu como *"quanto o Thiago
pagou"* — leitura óbvia, já que "quem pagou" estava na coluna ao lado.

Ele marcou 100%. O app entendeu "esta conta é inteirinha do Thiago, a Ana não
participa". Nessa leitura o que ele pagou não gera dívida e o que ela pagou vira
dívida cheia dele. **O acerto de um mês inteiro saiu invertido.**

A matemática sempre esteve certa (validada em 54 meses). O rótulo é que estava
ambíguo. A correção não foi mexer na conta: foi **eliminar o campo que exigia
adivinhar o significado** e trocar por um seletor em português de gente, mais uma
coluna **Resultado** que mostra, linha a linha e antes de gravar, quem fica
devendo quanto.

**Lição:** um campo que exige interpretação é um bug esperando acontecer, mesmo
com a aritmética perfeita.

---

## Estrutura da planilha

Cinco abas. A de dados é `Lançamentos`, uma tabela chamada `tbLancamentos`
com 15 colunas:

```
1 Data | 2 Ano | 3 Mês | 4 Competência | 5 Tipo | 6 Categoria | 7 Descrição
8 Parcela | 9 Valor | 10 Quem Pagou | 11 % Thiago | 12 Cota Thiago
13 Cota Ana | 14 Ana deve Thiago | 15 Observação
```

A coluna 11 continua se chamando `% Thiago` **no arquivo** (é o rateio, 0 a 1).
O que mudou foi a interface: o usuário nunca mais vê esse número, escolhe
"Meio a meio / Conta do Thiago / Conta da Ana".

Fórmulas das colunas calculadas:

```excel
Cota Thiago     = ROUND([@Valor]*[@[% Thiago]],2)
Cota Ana        = ROUND([@Valor]-[@[Cota Thiago]],2)
Ana deve Thiago = IF([@[Quem Pagou]]="Thiago",[@[Cota Ana]],
                  IF([@[Quem Pagou]]="Ana",-[@[Cota Thiago]],0))
```

Positivo = Ana deve ao Thiago. Negativo = Thiago deve à Ana.

Outras abas: `Painel` (resumo e 3 gráficos), `Config`, `Recorrentes`
(as despesas fixas, com liga/desliga), `Acertos` (as transferências feitas).

---

## Histórico

1. **Auditoria da planilha antiga.** 8 abas, 5 anos, 3.704 fórmulas, 3 layouts
   horizontais diferentes. Achados: Set–Dez/2026 eram cópia literal de 2025
   (~R$ 80 mil de despesa fantasma), linhas trocando de posição entre meses,
   contadores de parcela congelados, 49 variações de rótulo para as mesmas coisas.

2. **Migração.** 1.301 lançamentos extraídos e reconciliados contra os totais
   que a própria planilha antiga calculava. Layout novo, vertical, uma linha por
   lançamento.

3. **App local** (`file://`), lendo o xlsx com ZIP nativo
   (`DecompressionStream`/`CompressionStream`) e CRC32 próprio. Zero dependências.

4. **Auditoria própria.** Achei 4 bugs meus: faixa fixa de varredura, rateio
   assumido 50/50 quando 98 linhas não eram, "quem pagou" vazio virando Thiago,
   e um caractere de controle corrompendo o arquivo.

5. **Revisão externa.** 11 apontamentos, 8 reais. O mais grave:
   `Math.round(x*100)/100` divergia do Excel em 1 centavo.

6. **Migração para a nuvem.** O arquivo do Thiago vive só no OneDrive, sem
   sincronizar. Isso matou a File System Access API e forçou Microsoft Graph,
   com login PKCE escrito à mão (sem MSAL — o endpoint de token da Microsoft
   libera CORS, então dá).

7. **O bug do rateio invertido** e a correção descrita acima.

8. **Stress test completo** a pedido do Thiago: 8 baterias, 114 verificações,
   zero falha.

---

## A bateria de testes

Mora na pasta de trabalho da sessão, não no repositório. Roda com
`rodar-tudo.ps1` e não usa mock: carrega o `index.html` **real** dentro do jsdom,
para que o teste acompanhe o app quando ele mudar.

| Bateria | O que prova |
|---|---|
| 1. Regra central | A regra contra um oráculo em **centavos inteiros** escrito do zero; 40.000 casos aleatórios |
| 2. Interface | Tela de fixas e formulário, clicando de verdade no DOM |
| 3. Excel real | Abre no Excel via COM, recalcula e compara célula a célula |
| 4. Robustez | Concorrência, valores extremos, texto perigoso, ano inteiro |
| 5. Integridade | Os arquivos gravados abrem sem reparo |
| 6. Histórico | Os 1.301 lançamentos, 54 meses reconciliados |
| 7. Ciclo do mês | Lançar → acerto → transferir → zerar |
| 8. Remoção | Apagar lançamentos sem estragar o resto |

**Regra combinada com o Thiago:** nenhuma mudança é apresentada a ele sem a
bateria inteira passando com zero falha.

---

## Armadilhas que custaram caro

**`[@Coluna]` é sintaxe só de interface.** Gravada no XML gera `#REF!`. A forma
de arquivo é `tbLancamentos[[#This Row],[Coluna]]`. E `[@[Col Com Espaço]]`
corrompe a pasta inteira.

**Arredondamento.** `Math.round(9749.095*100)/100` dá `9749.09` porque o produto
em float vira `974909.4999...`. O Excel dá `9749.10`. Solução:

```js
function round2(x) {
  const s = Math.sign(x), a = Math.abs(x);
  return s * Number(Math.round(Number(a + "e+2")) + "e-2");
}
```

**Graph em OneDrive pessoal.** A API de Excel (`/workbook/`) **não suporta**
OneDrive pessoal, só business. Sobra baixar e subir o arquivo inteiro via
`/drives/{id}/items/{id}/content`, com `If-Match` no eTag para não sobrescrever
o que a outra pessoa gravou.

**Arredondar por linha desvia.** Cada linha de valor ímpar em centavos desvia
meio centavo. Nove linhas, quatro ímpares → 2 centavos de diferença contra
"metade do total". **Não é bug**: a planilha original faz igual, com `ROUND` na
fórmula. Manter.

**Caracteres de controle** (NUL, VT, ESC) corrompem o xlsx para o Excel, embora
o leitor JS aceite. Emoji, CJK, RTL e XML injetado passam sem problema.

**openpyxl** (usado para gerar a planilha): eixos de gráfico vêm com `delete=1`,
precisam de `delete=False`. Em `BarChart`, `x_axis` é sempre o eixo de categoria
mesmo com `type="bar"`. Validação de dados com `=` no `formula1` invalida o
arquivo; usar nome definido.

**Testar em jsdom:** o `Blob` dele não tem `.stream()` (injetar o do Node);
`const`/`let` de topo de script não viram propriedade de `window` (usar
`win.eval("NOME")`); e `Date` vindo do contexto do jsdom falha em
`instanceof Date` do Node.

---

## Arquivo travado pelo Excel

Gravar falha com `The resource you are attempting to access is locked` quando a
planilha está aberta no Excel Online ou no Excel desktop. Leitura continua
funcionando; só a escrita trava. Libera sozinho quando a sessão do Excel solta.

Aconteceu de verdade em 27/09/2026. A mensagem que o app mostra ainda é a do
Graph, em inglês — deveria dizer "A planilha está aberta no Excel. Feche e
tente de novo."

## Pontos em aberto

- **Ana gravando.** Ela lê, mas gravar em OneDrive pessoal de terceiro tem
  restrições que o corporativo não tem. Não testado — só a conta dela resolve.
- **Ruído de float no histórico.** ~400 lançamentos antigos com `0.499999` em
  vez de `0.5`, herdado da migração. Sem efeito prático (54 meses reconciliam em
  R$ 0,00). Limpeza oferecida, não autorizada.
- **`App Despesas.html` local** está desatualizado, sem a correção da Divisão.
  Decidir se mantém ou descarta.

---

*Última atualização: 27 de setembro de 2026.*

