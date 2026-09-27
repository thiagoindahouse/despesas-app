/* ============================================================================
   Remove linhas da aba Lancamentos de um xlsx, mantendo a pasta valida.
   Escrito como funcao pura para eu poder testar antes de rodar na planilha
   do Thiago: remover linha e mais perigoso que acrescentar, porque mexe na
   numeracao de tudo que vem depois.
   ========================================================================== */

/* Recebe o texto do sheet XML e os numeros de linha a remover.
   Devolve o XML novo e quantas linhas sobraram.                              */
function removerLinhasDoSheet(xml, remover) {
  const alvo = new Set(remover);
  const linhas = [];
  // captura cada <row ...>...</row> (e tambem <row .../> vazio)
  const re = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
  let m, ultimoFim = 0, prefixo = "", sufixo = "";
  const corpo = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml);
  if (!corpo) throw new Error("nao achei <sheetData> na aba");
  const dentro = corpo[1];
  while ((m = re.exec(dentro)) !== null) {
    const nr = /\br="(\d+)"/.exec(m[1]);
    linhas.push({ n: nr ? +nr[1] : 0, xml: m[0], attrs: m[1] });
  }
  const mantidas = linhas.filter(l => !alvo.has(l.n));
  const removidas = linhas.length - mantidas.length;

  /* Renumera: a linha 1 e o cabecalho e nunca sai. As demais recebem numeros
     sequenciais sem buraco, e cada celula tem que acompanhar (r="A15" -> "A12"). */
  let proximo = 0;
  const novas = mantidas.map(l => {
    proximo++;
    if (l.n === proximo) return l.xml;               // ja esta no lugar certo
    const de = l.n, para = proximo;
    let nx = l.xml.replace(/(<row\b[^>]*?\br=")(\d+)(")/, `$1${para}$3`);
    nx = nx.replace(/(<c\b[^>]*?\br="[A-Z]+)(\d+)(")/g, (_, a, num, c) =>
      (+num === de ? a + para + c : a + num + c));
    return nx;
  });

  let novo = xml.replace(/(<sheetData\b[^>]*>)[\s\S]*?(<\/sheetData>)/,
                         (_, a, b) => a + novas.join("") + b);
  const ultima = Math.max(1, novas.length);
  novo = novo.replace(/(dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/, `$1${ultima}$3`);
  return { xml: novo, ultima, removidas, total: novas.length };
}

module.exports = { removerLinhasDoSheet };
