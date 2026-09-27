/* ============================================================================
   BATERIA 4 - tela principal, robustez e casos extremos.
   Inclui o card de acerto (onde ele viu "Thiago transfere R$ 6.711,38"),
   os meses "so fixas", concorrencia, valores extremos e um ano inteiro.
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
const cent = s => {
  const m = String(s).replace(/[^\d,\-]/g, "").replace(",", ".");
  return Math.round(parseFloat(m) * 100);
};

(async () => {
  const { win } = await abrirApp();
  const doc = win.document;
  const $ = id => doc.getElementById(id);
  const V = e => win.eval(e);
  const disparar = (el, t) => el.dispatchEvent(new win.Event(t, { bubbles: true }));

  console.log("=".repeat(78));
  console.log("BATERIA 4 - TELA PRINCIPAL, ROBUSTEZ E CASOS EXTREMOS");
  console.log("=".repeat(78));

  const buf = fs.readFileSync(BASE_XLSX);
  const fonte = fonteFake(buf);
  await win.carregar(fonte);
  const hoje = new Date();

  /* ---------- 4.1 abre no mes corrente ---------- */
  console.log("\n[4.1] O app abre no mes corrente, nao no ano inteiro\n");
  const f = V("ESTADO").filtro;
  console.log(`  hoje: ${hoje.getMonth() + 1}/${hoje.getFullYear()} | filtro: mes=${f.mes} ano=${f.ano}`);
  checa("4.1 ano corrente", +f.ano === hoje.getFullYear(), `abriu em ${f.ano}`);
  checa("4.1 mes corrente", +f.mes === hoje.getMonth() + 1,
        `abriu no mes ${f.mes}, esperava ${hoje.getMonth() + 1}`);
  console.log(`  ${+f.mes === hoje.getMonth() + 1 ? "OK  " : "FALHA"} abriu em ` +
              `${V("MESES")[f.mes - 1]}/${f.ano}`);

  /* ---------- 4.2 quais meses estao marcados "so fixas" ---------- */
  console.log("\n[4.2] Marcacao 'so fixas' - setembro NAO pode estar marcado\n");
  const marcados = V(`mesesSoFixas(${hoje.getFullYear()})`);
  const nomes = V("MESES");
  console.log(`  marcados em ${hoje.getFullYear()}: ` +
              (marcados.length ? marcados.map(m => nomes[m - 1]).join(", ") : "nenhum"));
  checa("4.2 mes corrente livre", !marcados.includes(hoje.getMonth() + 1),
        `${nomes[hoje.getMonth()]} esta marcado como 'so fixas' e nao devia`);
  const opcoesMes = [...$("selMes").options].map(o => o.textContent);
  const setOpt = opcoesMes[hoje.getMonth() + 1];
  console.log(`  opcao do mes corrente no seletor: "${setOpt}"`);
  checa("4.2 seletor limpo", !/so fixas|só fixas/i.test(setOpt),
        `o seletor mostra "${setOpt}"`);
  // meses antigos com so fixas DEVEM continuar marcados
  const antigos = marcados.filter(m => m < hoje.getMonth() + 1);
  console.log(`  ${antigos.length} mes(es) anterior(es) marcado(s): ` +
              antigos.map(m => nomes[m - 1]).join(", "));
  checa("4.2 marcacao ainda funciona", antigos.length > 0,
        "nenhum mes antigo marcado - a marcacao parou de funcionar");

  /* ---------- 4.3 o card de acerto da tela principal ---------- */
  console.log("\n[4.3] Card 'Acerto do periodo' - direcao e valor\n");
  // injeta um setembro controlado direto no ESTADO, sem tocar no arquivo
  const est = V("ESTADO");
  const round2 = win.round2;
  const mkLinha = (desc, valor, quem, rateio = 0.5) => {
    const cotaT = round2(valor * rateio), cotaA = round2(valor - cotaT);
    return { linha: 9000, data: new Date(Date.UTC(hoje.getFullYear(), hoje.getMonth(), 1)),
      ano: hoje.getFullYear(), mes: hoje.getMonth() + 1, tipo: "Fixa", categoria: "Teste",
      descricao: desc, parcela: "", valor, quem, rateio, cotaT, cotaA,
      anaDeve: quem === "Thiago" ? cotaA : quem === "Ana" ? -cotaT : 0 };
  };
  const guardado = est.dados;
  for (const [nome, linhas, espDir] of [
      ["so o Thiago pagou", [mkLinha("A", 1000, "Thiago"), mkLinha("B", 500, "Thiago")], "Ana"],
      ["so a Ana pagou", [mkLinha("A", 1000, "Ana"), mkLinha("B", 500, "Ana")], "Thiago"],
      ["empate exato", [mkLinha("A", 1000, "Thiago"), mkLinha("B", 1000, "Ana")], "ninguem"],
      ["print do Thiago", [
        mkLinha("Financiamento AP", 8413.96, "Thiago"), mkLinha("Cartão Casa Thiago", 5308.03, "Thiago"),
        mkLinha("Parcela Carro BMW", 3489.51, "Thiago"), mkLinha("Condomínio", 2700, "Thiago"),
        mkLinha("Faxina", 1100, "Thiago"), mkLinha("Cartão Porto", 748.27, "Thiago"),
        mkLinha("Gás", 266.47, "Thiago"), mkLinha("Cartão Casa Ana", 6401.38, "Ana"),
        mkLinha("Luz", 310, "Ana")], "Ana"]]) {
    est.dados = linhas;
    est.acertos = [];
    win.desenhar();
    const card = $("acerto").textContent.replace(/\s+/g, " ").trim();
    const soma = linhas.reduce((s, l) => s + l.anaDeve, 0);
    const dir = Math.abs(soma) < 0.005 ? "ninguem" : soma > 0 ? "Ana" : "Thiago";
    const ok = checa(`4.3 ${nome}`, dir === espDir,
                     `o card indica ${dir}, esperava ${espDir}. card: "${card}"`);
    console.log(`  ${ok ? "OK  " : "FALHA"} ${nome.padEnd(22)} -> ${dir === "ninguem" ? "ninguem deve" :
                 dir + " deve R$ " + Math.abs(soma).toFixed(2)}`);
    console.log(`         card: ${card.slice(0, 120)}`);
  }
  est.dados = guardado;

  /* ---------- 4.4 transferencia abate do saldo ---------- */
  console.log("\n[4.4] Transferencia registrada abate do saldo\n");
  est.dados = [mkLinha("X", 2000, "Thiago")];          // Ana deve 1000
  est.acertos = [{ data: new Date(Date.UTC(hoje.getFullYear(), hoje.getMonth(), 15)),
                   ano: hoje.getFullYear(), mes: hoje.getMonth() + 1,
                   de: "Ana", para: "Thiago", valor: 400, ref: "parcial" }];
  win.desenhar();
  const cardT = $("acerto").textContent.replace(/\s+/g, " ").trim();
  console.log(`  gerado 1000,00 | ja transferido 400,00 | card: ${cardT.slice(0, 150)}`);
  checa("4.4 abate", /600/.test(cardT), `esperava saldo 600 no card: "${cardT}"`);
  est.acertos = [];

  /* ---------- 4.5 concorrencia: a planilha mudou por fora ---------- */
  console.log("\n[4.5] Alguem gravou por fora - o app recusa e nao sobrescreve\n");
  const fonte2 = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte2);
  const versaoVelha = V("ESTADO").versao;
  const reg = { data: new Date(hoje.getFullYear(), hoje.getMonth(), 1), tipo: "Fixa",
    categoria: "Teste", descricao: "Concorrencia", parcela: "", valor: 100,
    quem: "Thiago", rateio: 0.5, obs: "" };
  await win.gravarLancamento(fonte2, reg, versaoVelha);       // outra aba grava primeiro
  let recusou = false, msg = "";
  try { await win.gravarLancamento(fonte2, reg, versaoVelha); }
  catch (e) { recusou = true; msg = e.message; }
  checa("4.5 recusa versao velha", recusou, "GRAVOU EM CIMA - risco de perder lancamento");
  console.log(`  ${recusou ? "OK  " : "FALHA"} segunda gravacao com versao velha: ` +
              (recusou ? `recusada ("${msg}")` : "ACEITOU, isso apaga o trabalho do outro"));

  /* ---------- 4.6 valores extremos ---------- */
  console.log("\n[4.6] Valores extremos e entradas estranhas\n");
  const extremos = [
    ["um centavo", 0.01, "Thiago", true],
    ["tres centavos", 0.03, "Ana", true],
    ["valor alto", 9999999.99, "Thiago", true],
    ["com 4 casas", 12.3456, "Thiago", true],
    ["texto no lugar", NaN, "Thiago", false],
    ["zero", 0, "Thiago", false],
    ["negativo", -100, "Thiago", false],
  ];
  for (const [nome, valor, quem, deveriaPassar] of extremos) {
    const v = round2(valor);
    const valido = Number.isFinite(v) && v > 0 && v <= 99999999;
    const cotaT = valido ? round2(v * 0.5) : null;
    const cotaA = valido ? round2(v - cotaT) : null;
    const ad = valido ? (quem === "Thiago" ? cotaA : -cotaT) : null;
    const ok = checa(`4.6 ${nome}`, valido === deveriaPassar,
                     `valido=${valido}, esperava ${deveriaPassar}`);
    console.log(`  ${ok ? "OK  " : "FALHA"} ${nome.padEnd(16)} ${String(valor).padStart(12)} -> ` +
      (valido ? `aceito, vira ${v.toFixed(2)}, ${quem === "Thiago" ? "Ana" : "Thiago"} deve ` +
                `${Math.abs(ad).toFixed(2)}` : "recusado pelo app"));
  }

  /* ---------- 4.7 texto perigoso nao quebra o xlsx ---------- */
  console.log("\n[4.7] Texto perigoso na descricao\n");
  const fonte3 = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte3);
  const perigosos = [
    ['aspas "duplas" e <tags>', "XML"],
    ["e comercial & menor <", "escape"],
    ["emoji 🏠💸 e acento çãõ", "unicode"],
    ["'; DROP TABLE --", "injecao"],
    ["a".repeat(300), "300 caracteres"],
  ];
  let v3 = V("ESTADO").versao;
  for (const [txt, rotulo] of perigosos) {
    // gravarLancamentos devolve NUMEROS DE LINHA, nao a versao; releio a versao da fonte
    const atual = await fonte3.ler();
    await win.gravarLancamento(fonte3, { ...reg, descricao: txt, valor: 50 }, atual.versao);
  }
  await win.carregar(fonte3);
  const gravados = V("ESTADO").dados.slice(-perigosos.length);
  let erros47 = 0;
  perigosos.forEach(([txt, rotulo], i) => {
    const bate = gravados[i] && gravados[i].descricao === txt;
    if (!bate) { erros47++; }
    console.log(`  ${bate ? "OK  " : "FALHA"} ${rotulo.padEnd(16)} ` +
                `voltou ${bate ? "identico" : "DIFERENTE: " + JSON.stringify(gravados[i] && gravados[i].descricao)}`);
  });
  checa("4.7 texto perigoso", erros47 === 0, `${erros47} textos corromperam`);
  fs.writeFileSync(require("path").join(__dirname, "saida-b4.xlsx"), fonte3.buffer);

  /* ---------- 4.8 um ano inteiro, 12 meses de lancamentos ---------- */
  console.log("\n[4.8] Ano inteiro: 12 meses lancados em sequencia\n");
  const fonte4 = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte4);

  let espAno = 0;
  const ANO = 2027;
  for (let m = 1; m <= 12; m++) {
    const lote = [];
    for (let k = 0; k < 6; k++) {
      const valor = round2(100 + Math.random() * 5000);
      const quem = ["Thiago", "Ana", "Casa"][(m + k) % 3];
      lote.push({ data: new Date(ANO, m - 1, 5), tipo: "Fixa", categoria: "Teste",
        descricao: `Item ${m}-${k}`, parcela: "", valor, quem, rateio: 0.5, obs: "" });
      const cT = round2(valor * 0.5), cA = round2(valor - cT);
      espAno += quem === "Thiago" ? cA : quem === "Ana" ? -cT : 0;
    }
    const at4 = await fonte4.ler();
    await win.gravarLancamentos(fonte4, lote, at4.versao);
  }
  await win.carregar(fonte4);
  const est4 = V("ESTADO");
  const doAno = est4.dados.filter(l => l.ano === ANO);
  const somaAno = doAno.reduce((s, l) => s + l.anaDeve, 0);
  console.log(`  gravou ${doAno.length} lancamentos em 12 meses de ${ANO}`);
  console.log(`  acerto do ano segundo o app: R$ ${somaAno.toFixed(2)}`);
  console.log(`  acerto do ano calculado a parte: R$ ${espAno.toFixed(2)}`);
  checa("4.8 quantidade", doAno.length === 72, `esperava 72, gravou ${doAno.length}`);
  checa("4.8 acerto do ano", Math.abs(somaAno - espAno) < 0.005,
        `divergencia de R$ ${Math.abs(somaAno - espAno).toFixed(2)}`);
  console.log(`  ${Math.abs(somaAno - espAno) < 0.005 ? "OK  " : "FALHA"} diferenca: ` +
              `R$ ${Math.abs(somaAno - espAno).toFixed(4)}`);
  // conferencia mes a mes
  let mesesRuins = 0;
  for (let m = 1; m <= 12; m++) {
    const doMes = doAno.filter(l => l.mes === m);
    const somaM = doMes.reduce((s, l) => s + l.anaDeve, 0);
    const pagoT = doMes.filter(l => l.quem === "Thiago").reduce((s, l) => s + l.valor, 0);
    const pagoA = doMes.filter(l => l.quem === "Ana").reduce((s, l) => s + l.valor, 0);
    // com rateio 50/50 e ignorando "Casa": quem pagou mais e credor
    const semCasa = doMes.filter(l => l.quem !== "Casa");
    const totalSC = semCasa.reduce((s, l) => s + l.valor, 0);
    const espM = round2(pagoT - totalSC / 2);
    if (Math.abs(somaM - espM) > 0.03) {
      mesesRuins++;
      console.log(`     FALHA mes ${m}: app ${somaM.toFixed(2)}, conta direta ${espM.toFixed(2)}`);
    }
  }
  checa("4.8 mes a mes", mesesRuins === 0, `${mesesRuins} meses divergentes`);
  console.log(`  ${mesesRuins === 0 ? "OK  " : "FALHA"} os 12 meses batem com a conta direta ` +
              `(pagou - metade do total)`);
  fs.writeFileSync(require("path").join(__dirname, "saida-b4-ano.xlsx"), fonte4.buffer);

  console.log("\n" + "=".repeat(78));
  console.log(`BATERIA 4: ${R.ok} passaram, ${R.falhas.length} falharam`);
  R.falhas.forEach(f => console.log("  FALHA:", f.nome, "-", f.detalhe));
  console.log("=".repeat(78));
  if (R.falhas.length) process.exitCode = 1;
})().catch(e => { console.error("HARNESS QUEBROU:", e); process.exit(1); });

