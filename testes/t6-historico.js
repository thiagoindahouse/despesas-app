/* ============================================================================
   BATERIA 6 - reconciliacao historica e o que faltou.
   Passa os 1.301 lancamentos reais pela regra e confere mes a mes contra a
   conta direta: quem pagou mais e credor da diferenca.
   ========================================================================== */
const { abrirApp, fonteFake, BASE_XLSX } = require("./harness");
const fs = require("fs");

const R = { ok: 0, falhas: [] };
function checa(nome, cond, detalhe) {
  if (cond) { R.ok++; return true; }
  R.falhas.push({ nome, detalhe });
  console.log(`  FALHA ${nome}: ${detalhe}`);
  return false;
}

(async () => {
  const { win } = await abrirApp();
  const $ = id => win.document.getElementById(id);
  const V = e => win.eval(e);
  const round2 = win.round2;

  console.log("=".repeat(78));
  console.log("BATERIA 6 - RECONCILIACAO HISTORICA");
  console.log("=".repeat(78));

  const fonte = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte);
  const dados = V("ESTADO").dados;
  console.log(`\n  ${dados.length} lancamentos, de ${Math.min(...dados.map(l => l.ano))} ` +
              `a ${Math.max(...dados.map(l => l.ano))}`);

  /* ---------- 6.1 a direcao do acerto em TODOS os meses ---------- */
  console.log("\n[6.1] Direcao do acerto em todos os meses da planilha\n");
  const meses = [...new Set(dados.map(l => `${l.ano}-${String(l.mes).padStart(2, "0")}`))].sort();
  let ruins = 0, maiorDif = 0, piorMes = null, somaDifs = 0;
  const linhasTab = [];
  for (const ym of meses) {
    const [a, m] = ym.split("-").map(Number);
    const doMes = dados.filter(l => l.ano === a && l.mes === m);
    const app = doMes.reduce((s, l) => s + l.anaDeve, 0);

    // conta direta, independente: so o que tem rateio 50/50 e pagador definido
    const rel = doMes.filter(l => l.quem === "Thiago" || l.quem === "Ana");
    const pagoT = rel.filter(l => l.quem === "Thiago").reduce((s, l) => s + l.valor, 0);
    const cabeT = rel.reduce((s, l) => s + round2(l.valor * l.rateio), 0);
    const direta = pagoT - cabeT;     // pagou a mais do que cabia a ele

    const dif = Math.abs(app - direta);
    somaDifs += dif;
    const impares = rel.filter(l => Math.round(l.valor * 100) % 2 !== 0).length;
    const tol = Math.max(0.02, impares * 0.01);
    if (dif > tol) {
      ruins++;
      if (dif > maiorDif) { maiorDif = dif; piorMes = { ym, app, direta, dif }; }
    }
    linhasTab.push({ ym, n: doMes.length, app, direta, dif });
  }
  console.log(`  ${meses.length} meses conferidos`);
  console.log(`  ${ruins} meses fora da tolerancia de arredondamento`);
  console.log(`  soma de todas as diferencas nos ${meses.length} meses: R$ ${somaDifs.toFixed(2)}`);
  if (piorMes) console.log(`  pior mes: ${piorMes.ym} app=${piorMes.app.toFixed(2)} ` +
                           `direta=${piorMes.direta.toFixed(2)} dif=${piorMes.dif.toFixed(2)}`);
  checa("6.1 reconciliacao mes a mes", ruins === 0, `${ruins} meses divergentes`);

  // amostra dos 12 meses mais recentes
  console.log("\n  ultimos 12 meses:");
  console.log("     mes      n   acerto do app   conta direta   dif");
  linhasTab.slice(-12).forEach(l =>
    console.log(`     ${l.ym}  ${String(l.n).padStart(3)}  ${l.app.toFixed(2).padStart(13)}  ` +
                `${l.direta.toFixed(2).padStart(13)}  ${l.dif.toFixed(2).padStart(5)}`));

  /* ---------- 6.2 nenhum lancamento com direcao impossivel ---------- */
  console.log("\n[6.2] Nenhum lancamento com direcao impossivel\n");
  const impossiveis = dados.filter(l => {
    if (l.quem === "Thiago" && l.rateio < 0.999 && !(l.anaDeve > 0)) return true;
    if (l.quem === "Ana" && l.rateio > 0.001 && !(l.anaDeve < 0)) return true;
    if (l.quem !== "Thiago" && l.quem !== "Ana" && Math.abs(l.anaDeve) > 0.005) return true;
    return false;
  });
  checa("6.2 direcao coerente", impossiveis.length === 0,
        `${impossiveis.length} lancamentos com direcao impossivel`);
  console.log(`  ${impossiveis.length === 0 ? "OK  " : "FALHA"} ` +
              `${dados.length} lancamentos, ${impossiveis.length} com direcao impossivel`);
  if (impossiveis.length) impossiveis.slice(0, 5).forEach(l =>
    console.log("     ", JSON.stringify({ l: l.linha, d: l.descricao, q: l.quem, r: l.rateio, ad: l.anaDeve })));

  /* ---------- 6.3 distribuicao dos rateios no historico ---------- */
  console.log("\n[6.3] Como os rateios estao distribuidos no historico\n");
  const porRateio = new Map();
  dados.forEach(l => porRateio.set(l.rateio, (porRateio.get(l.rateio) || 0) + 1));
  [...porRateio.entries()].sort((a, b) => b[1] - a[1]).forEach(([r, n]) => {
    const rot = Math.abs(r - 0.5) < 0.005 ? "meio a meio"
              : r >= 0.995 ? "conta do Thiago" : r <= 0.005 ? "conta da Ana" : `${(r*100).toFixed(1)}% Thiago`;
    console.log(`     ${String(n).padStart(5)} lancamentos  rateio ${r}  (${rot})`);
  });
  const foraDaFaixa = dados.filter(l => l.rateio < 0 || l.rateio > 1);
  checa("6.3 rateio na faixa", foraDaFaixa.length === 0, `${foraDaFaixa.length} fora de 0..1`);

  /* ---------- 6.4 gravar uma transferencia na aba Acertos ---------- */
  console.log("\n[6.4] Registrar transferencia na aba Acertos\n");
  const antes = V("ESTADO").acertos.length;
  const at = await fonte.ler();
  await win.gravarAcerto(fonte, {
    data: new Date(2026, 8, 30), de: "Ana", para: "Thiago",
    valor: 7657.41, ref: "acerto de setembro/2026 (teste)",
  }, at.versao);
  await win.carregar(fonte);
  const depois = V("ESTADO").acertos;
  console.log(`  acertos: ${antes} -> ${depois.length}`);
  checa("6.4 gravou", depois.length === antes + 1, `nao gravou a transferencia`);
  const ultimo = depois[depois.length - 1];
  console.log(`  ultimo: ${JSON.stringify({ de: ultimo.de, para: ultimo.para, valor: ultimo.valor,
                                            ano: ultimo.ano, mes: ultimo.mes })}`);
  checa("6.4 valores certos", ultimo.de === "Ana" && ultimo.para === "Thiago" &&
        Math.abs(ultimo.valor - 7657.41) < 0.005, "a transferencia voltou diferente");
  fs.writeFileSync(require("path").join(__dirname, "saida-b6.xlsx"), fonte.buffer);

  /* ---------- 6.5 duplicata: o app avisa? ---------- */
  console.log("\n[6.5] Deteccao de fixa ja lancada no mes\n");
  win.abrirFixas();
  const hoje = new Date();
  $("fxMes").value = `${hoje.getFullYear()}-${hoje.getMonth() + 1}`;
  win.montarFixas();
  const chips = [...$("fxCorpo").querySelectorAll(".chip")].map(c => c.textContent.trim());
  const marcadasFora = [...$("fxCorpo").querySelectorAll("tr[data-i]")]
    .filter(tr => !tr.querySelector(".inc").checked).length;
  console.log(`  chips na tela: ${JSON.stringify([...new Set(chips)])}`);
  console.log(`  linhas ja vindo desmarcadas: ${marcadasFora}`);
  const fixasDoMes = V(`fixasDoMes(${hoje.getFullYear()}, ${hoje.getMonth() + 1})`);
  const dup = fixasDoMes.filter(i => i.duplicado).length;
  console.log(`  o app marcou ${dup} como ja lancada(s)`);
  checa("6.5 duplicata coerente", dup === marcadasFora,
        `${dup} duplicadas mas ${marcadasFora} desmarcadas`);

  /* ---------- 6.6 BMW e Financiamento aparecem ---------- */
  console.log("\n[6.6] As fixas que ele cobrou estao na lista?\n");
  const nomesFixas = fixasDoMes.map(i => i.descricao);
  for (const alvo of ["Parcela Carro BMW", "Financiamento AP", "Cartão Casa Thiago",
                      "Cartão Casa Ana", "Cartão Porto", "Condomínio", "Faxina", "Gás", "Luz"]) {
    const tem = nomesFixas.includes(alvo);
    checa(`6.6 ${alvo}`, tem, "nao aparece nas fixas do mes");
    console.log(`  ${tem ? "OK  " : "FALHA"} ${alvo}`);
  }
  for (const fora of ["Seguro Casa", "Apple", "Globo"]) {
    const tem = nomesFixas.some(n => n.includes(fora));
    checa(`6.6 sem ${fora}`, !tem, "ainda aparece e devia estar desligada");
    console.log(`  ${!tem ? "OK  " : "FALHA"} ${fora} fora da lista (duplicava com o cartao)`);
  }

  console.log("\n" + "=".repeat(78));
  console.log(`BATERIA 6: ${R.ok} passaram, ${R.falhas.length} falharam`);
  R.falhas.forEach(f => console.log("  FALHA:", f.nome, "-", f.detalhe));
  console.log("=".repeat(78));
  if (R.falhas.length) process.exitCode = 1;
})().catch(e => { console.error("HARNESS QUEBROU:", e); process.exit(1); });
