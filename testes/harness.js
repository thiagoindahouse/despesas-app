/* Harness: carrega o index.html REAL dentro do jsdom e expoe as funcoes internas
   para teste. Nao reimplementa nada do app: se o app mudar, o teste acompanha.  */
const fs = require("fs");
const path = require("path");
const { JSDOM, VirtualConsole } = require("jsdom");

const APP = path.join(__dirname, "despesas-app", "index.html");
const BASE_XLSX = path.join(__dirname, "base-teste.xlsx");

/* O app usa DecompressionStream/CompressionStream, que o Node 24 tem em
   node:stream/web mas nao expoe no escopo do jsdom. Injeto os do Node.        */
function instalarStreams(win) {
  const w = require("node:stream/web");
  for (const k of ["CompressionStream", "DecompressionStream", "ReadableStream",
                   "WritableStream", "TransformStream"]) {
    if (!win[k] && w[k]) win[k] = w[k];
  }
  if (!win.Response) win.Response = globalThis.Response;
  if (!win.Blob) win.Blob = globalThis.Blob;
  if (!win.crypto) win.crypto = globalThis.crypto;
  /* O Blob do jsdom nao implementa .stream(), e o app depende disso para o ZIP.
     Uso os do Node, que sao os mesmos que o navegador expoe de verdade.        */
  if (typeof win.Blob.prototype.stream !== "function") win.Blob = globalThis.Blob;
  win.Response = globalThis.Response;
  if (!win.matchMedia) {
    win.matchMedia = q => ({
      matches: false, media: q, onchange: null,
      addEventListener() {}, removeEventListener() {},
      addListener() {}, removeListener() {}, dispatchEvent() { return false; },
    });
  }
}

async function abrirApp(opts = {}) {
  const html = fs.readFileSync(APP, "utf8");
  const vc = new VirtualConsole();
  const erros = [];
  vc.on("jsdomError", e => erros.push(e));
  vc.on("error", (...a) => erros.push(new Error(a.join(" "))));

  const dom = new JSDOM(html, {
    runScripts: "dangerously",
    url: opts.url || "https://thiagoindahouse.github.io/despesas-app/",
    virtualConsole: vc,
    pretendToBeVisual: true,
    beforeParse(win) { instalarStreams(win); },
  });
  const win = dom.window;
  instalarStreams(win);

  // o app pode registrar coisas no load
  await new Promise(r => {
    if (win.document.readyState === "complete") return r();
    win.addEventListener("load", r);
  });
  await new Promise(r => setTimeout(r, 30));

  win.confirm = () => true;
  win.alert = () => {};
  if (!win.navigator.locks) {
    Object.defineProperty(win.navigator, "locks", {
      value: { request: async (_n, fn) => fn() }, configurable: true,
    });
  }
  return { dom, win, erros };
}

/* Fonte falsa: mesma interface que o app espera das fontes real (local/onedrive),
   mas guardando o buffer em memoria para eu poder inspecionar o que foi gravado. */
function fonteFake(buf, nome = "teste.xlsx") {
  let atual = Buffer.from(buf);
  let versao = 1;
  return {
    tipo: "teste", nome,
    gravacoes: [],
    get buffer() { return atual; },
    async ler() { return { buf: atual.buffer.slice(atual.byteOffset, atual.byteOffset + atual.length), versao }; },
    async gravar(novo, versaoEsperada) {
      if (versaoEsperada != null && versaoEsperada !== versao) {
        throw new Error("a planilha mudou desde que você abriu");
      }
      // o app entrega um Blob (e o Graph aceita Blob); aqui preciso de Buffer
      const bytes = (novo && typeof novo.arrayBuffer === "function")
        ? Buffer.from(await novo.arrayBuffer())
        : Buffer.from(novo);
      atual = bytes;
      versao++;
      this.gravacoes.push(atual.length);
      return versao;
    },
    async permissao() { return true; },
  };
}

module.exports = { abrirApp, fonteFake, APP, BASE_XLSX };
