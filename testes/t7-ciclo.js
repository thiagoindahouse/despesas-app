/* ============================================================================
   BATERIA 7 - ciclo completo de um mes, do jeito que o Thiago vai usar:
   lancar fixas -> ver o acerto -> registrar a transferencia -> conferir que
   o saldo zera. Se a data da transferencia nao voltar certa, o app cobra de
   novo um dinheiro que ja foi pago.
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
  const disparar = (el, t) => el.dispatchEvent(new win.Event(t, { bubbles: true }));

  console.log("=".repeat(78));
  console.log("BATERIA 7 - CICLO COMPLETO DO MES");
  console.log("=".repeat(78));

  const fonte = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte);
  const hoje = new Date();
  const ANO = hoje.getFullYear(), MES = hoje.getMonth() + 1;

  /* ---------- passo 1: lancar as fixas do mes ---------- */
  console.log(`\n[passo 1] Lancar as fixas de ${V("MESES")[MES-1]}/${ANO}\n`);
  win.abrirFixas();
  $("fxMes").value = `${ANO}-${MES}`;
  win.montarFixas();
  const cenario = {
    "Financiamento AP": [8413.96, "Thiago"], "Cartão Casa Thiago": [5308.03, "Thiago"],
    "Parcela Carro BMW": [3489.51, "Thiago"], "Condomínio": [2700, "Thiago"],
    "Faxina": [1100, "Thiago"], "Cartão Porto": [748.27, "Thiago"],
    "Gás": [266.47, "Thiago"], "Cartão Casa Ana": [6401.38, "Ana"], "Luz": [310, "Ana"],
  };
  const trs = [...$("fxCorpo").querySelectorAll("tr[data-i]")];
  trs.forEach(tr => {
    const d = tr.querySelector("td:nth-child(2)").textContent.trim().split("\n")[0].trim();
    const c = cenario[d];
    tr.querySelector(".inc").checked = !!c;
    if (c) {
      tr.querySelector(".val").value = c[0].toFixed(2);
      tr.querySelector(".quem").value = c[1];
      tr.querySelector(".div").value = "meio";
    }
  });
  disparar(trs[0].querySelector(".val"), "input");
  const rodape = $("fxResumo").textContent.replace(/\s+/g, " ").trim();
  console.log(`  ${rodape}`);
  checa("7.1 rodape diz Ana deve", /Ana deve/.test(rodape), `rodape: "${rodape}"`);

  const regs = win.lerFixasDaTela();
  const at1 = await fonte.ler();
  await win.gravarLancamentos(fonte, regs, at1.versao);
  await win.carregar(fonte);
  console.log(`  gravou ${regs.length} lancamentos`);

  /* ---------- passo 2: o card da tela principal ---------- */
  console.log("\n[passo 2] Card de acerto da tela principal\n");
  const est = V("ESTADO");
  est.filtro.ano = ANO; est.filtro.mes = MES; est.filtro.cat = ""; est.filtro.quem = "";
  win.desenhar();
  let card = $("acerto").textContent.replace(/\s+/g, " ").trim();
  console.log(`  ${card}`);
  const doMes = est.dados.filter(l => l.ano === ANO && l.mes === MES);
  const saldo = doMes.reduce((s, l) => s + l.anaDeve, 0);
  console.log(`\n  saldo calculado: R$ ${saldo.toFixed(2)} -> ${saldo > 0 ? "ANA deve" : "THIAGO deve"}`);
  checa("7.2 direcao no card", saldo > 0 && /Ana transfere/.test(card),
        `saldo ${saldo.toFixed(2)}, card "${card}"`);
  checa("7.2 sem transferencia ainda", !/Já transferido/i.test(card),
        "o card ja fala em transferencia antes de existir uma");

  /* ---------- passo 3: registrar a transferencia ---------- */
  console.log("\n[passo 3] A Ana transfere o valor - o app tem que zerar\n");
  const valorTransf = Math.round(saldo * 100) / 100;
  const at2 = await fonte.ler();
  await win.gravarAcerto(fonte, {
    data: new Date(ANO, MES - 1, 28), de: "Ana", para: "Thiago",
    valor: valorTransf, ref: `acerto de ${V("MESES")[MES-1]}/${ANO}`,
  }, at2.versao);
  await win.carregar(fonte);

  const acertos = V("ESTADO").acertos;
  const ult = acertos[acertos.length - 1];
  console.log(`  transferencia lida de volta:`);
  console.log(`     data : ${ult.data ? ult.data.toISOString().slice(0,10) : "NULL <<< PROBLEMA"}`);
  console.log(`     de   : ${ult.de} -> para: ${ult.para}`);
  console.log(`     valor: R$ ${ult.valor.toFixed(2)}`);
  checa("7.3 data preservada", ult.data && typeof ult.data.getUTCFullYear === "function" && !isNaN(ult.data.getTime()),
        "a data voltou nula; o app nao vai reconhecer a transferencia no periodo");
  if (ult.data) {
    const okAno = ult.data.getUTCFullYear() === ANO, okMes = ult.data.getUTCMonth() + 1 === MES;
    checa("7.3 periodo certo", okAno && okMes,
          `voltou ${ult.data.getUTCFullYear()}-${ult.data.getUTCMonth()+1}, esperava ${ANO}-${MES}`);
    console.log(`  ${okAno && okMes ? "OK  " : "FALHA"} a transferencia caiu no periodo certo`);
  }

  /* ---------- passo 4: o saldo tem que estar zerado ---------- */
  console.log("\n[passo 4] O card depois da transferencia\n");
  const est4 = V("ESTADO");
  est4.filtro.ano = ANO; est4.filtro.mes = MES;
  win.desenhar();
  card = $("acerto").textContent.replace(/\s+/g, " ").trim();
  console.log(`  ${card}`);
  const reconhece = /Já transferido/i.test(card);
  checa("7.4 reconhece a transferencia", reconhece,
        "o card ignorou a transferencia e vai cobrar de novo");
  const zerou = /tudo (acertado|equilibrado)|R\$ 0,00/i.test(card);
  checa("7.4 saldo zerado", zerou, `o saldo nao zerou. card: "${card}"`);
  console.log(`  ${reconhece ? "OK  " : "FALHA"} reconheceu a transferencia`);
  console.log(`  ${zerou ? "OK  " : "FALHA"} saldo em aberto zerado`);

  /* ---------- passo 5: transferencia no sentido contrario ---------- */
  console.log("\n[passo 5] Mes em que o Thiago e que deve\n");
  const fonte5 = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte5);
  const MES5 = MES;
  const lote5 = [
    { data: new Date(ANO, MES5 - 1, 3), tipo: "Fixa", categoria: "Teste", descricao: "So da Ana 1",
      parcela: "", valor: 5000, quem: "Ana", rateio: 0.5, obs: "" },
    { data: new Date(ANO, MES5 - 1, 4), tipo: "Fixa", categoria: "Teste", descricao: "So da Ana 2",
      parcela: "", valor: 3000, quem: "Ana", rateio: 0.5, obs: "" },
  ];
  const at5 = await fonte5.ler();
  await win.gravarLancamentos(fonte5, lote5, at5.versao);
  await win.carregar(fonte5);
  const est5 = V("ESTADO");
  est5.filtro.ano = ANO; est5.filtro.mes = MES5; est5.filtro.cat = ""; est5.filtro.quem = "";
  est5.dados = est5.dados.filter(l => !(l.ano === ANO && l.mes === MES5) ||
                                      l.descricao.startsWith("So da Ana"));
  win.desenhar();
  const card5 = $("acerto").textContent.replace(/\s+/g, " ").trim();
  console.log(`  Ana pagou R$ 8.000,00 sozinha`);
  console.log(`  ${card5}`);
  checa("7.5 direcao invertida", /Thiago transfere/.test(card5),
        `esperava "Thiago transfere", card: "${card5}"`);
  const esp5 = 4000;
  checa("7.5 valor", new RegExp(esp5.toLocaleString("pt-BR", {minimumFractionDigits:2})).test(card5),
        `esperava R$ 4.000,00 no card`);
  console.log(`  ${/Thiago transfere/.test(card5) ? "OK  " : "FALHA"} ` +
              `Thiago deve R$ 4.000,00 (metade dos 8.000 que a Ana pagou)`);

  fs.writeFileSync(require("path").join(__dirname, "saida-b7.xlsx"), fonte.buffer);

  console.log("\n" + "=".repeat(78));
  console.log(`BATERIA 7: ${R.ok} passaram, ${R.falhas.length} falharam`);
  R.falhas.forEach(f => console.log("  FALHA:", f.nome, "-", f.detalhe));
  console.log("=".repeat(78));
  if (R.falhas.length) process.exitCode = 1;
})().catch(e => { console.error("HARNESS QUEBROU:", e); process.exit(1); });

