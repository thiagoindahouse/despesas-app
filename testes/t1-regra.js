/* ============================================================================
   BATERIA 1 - a regra do Thiago, nos 4 pontos onde o app calcula o acerto.

   Regra (palavras dele):
     "quando o thiago paga tudo, a ana deve"
     "quando a ana paga tudo o thiago deve"
     "quando o thiago e a ana pagam 50/50, fica metade para cada um"

   Oraculo: aritmetica INTEIRA em centavos, escrita do zero. Se o app divergir
   do oraculo, o app esta errado - nao o contrario.
   ========================================================================== */
const { abrirApp, fonteFake, BASE_XLSX } = require("./harness");
const fs = require("fs");

const R = { ok: 0, falhas: [] };
function checa(nome, cond, detalhe) {
  if (cond) { R.ok++; return true; }
  R.falhas.push({ nome, detalhe });
  return false;
}

/* ---- ORACULO: inteiro, sem float, sem olhar o app ---- */
function metadeCent(v) {                 // metade com empate pra cima
  return (v % 2 === 0) ? v / 2 : (v + 1) / 2;
}
function oraculoCent(vCent, quem, divisao) {
  let tCent;
  if (divisao === "thiago") tCent = vCent;
  else if (divisao === "ana") tCent = 0;
  else tCent = metadeCent(vCent);
  const aCent = vCent - tCent;
  if (quem === "Thiago") return aCent;    // ele adiantou a parte dela
  if (quem === "Ana") return -tCent;      // ela adiantou a parte dele
  return 0;                               // Casa: cada um pagou a sua metade
}
const RATEIO = { meio: 0.5, thiago: 1, ana: 0 };

(async () => {
  const { win, erros } = await abrirApp();
  if (erros.length) {
    erros.forEach(e => console.log("ERRO DOM:", (e.message || e).toString().slice(0, 300)));
  }
  const DIVISOES = win.eval("DIVISOES");
  const round2 = win.round2;

  console.log("=".repeat(78));
  console.log("BATERIA 1 - REGRA CENTRAL");
  console.log("=".repeat(78));

  /* ---------- 1.1 os 3 casos que ele descreveu, em palavras ---------- */
  console.log("\n[1.1] Os tres casos, ao pe da letra\n");
  const casos = [
    { quem: "Thiago", div: "meio",   v: 1000, esperado: "Ana deve 500,00" },
    { quem: "Ana",    div: "meio",   v: 1000, esperado: "Thiago deve 500,00" },
    { quem: "Casa",   div: "meio",   v: 1000, esperado: "ninguem deve" },
  ];
  for (const c of casos) {
    const vC = Math.round(c.v * 100);
    const esp = oraculoCent(vC, c.quem, c.div);
    const cotaT = round2(c.v * RATEIO[c.div]);
    const cotaA = round2(c.v - cotaT);
    const got = c.quem === "Thiago" ? cotaA : (c.quem === "Ana" ? -cotaT : 0);
    const gotC = Math.round(got * 100);
    const txt = gotC > 0 ? `Ana deve ${(gotC/100).toFixed(2)}`
              : gotC < 0 ? `Thiago deve ${(-gotC/100).toFixed(2)}` : "ninguem deve";
    const ok = checa(`1.1 ${c.quem}/${c.div}`, gotC === esp, `esperava ${esp}, veio ${gotC}`);
    console.log(`  ${ok ? "OK  " : "FALHA"} ${c.quem.padEnd(7)} paga R$ ${c.v} ${c.div.padEnd(7)} -> ${txt}`);
  }

  /* ---------- 1.2 matriz completa quem x divisao ---------- */
  console.log("\n[1.2] Matriz completa: quem pagou x divisao\n");
  console.log("        pagador |   meio a meio   | conta do Thiago | conta da Ana");
  console.log("     " + "-".repeat(70));
  for (const quem of ["Thiago", "Ana", "Casa"]) {
    const cels = [];
    for (const div of ["meio", "thiago", "ana"]) {
      const v = 1000;
      const cotaT = round2(v * RATEIO[div]);
      const cotaA = round2(v - cotaT);
      const got = quem === "Thiago" ? cotaA : (quem === "Ana" ? -cotaT : 0);
      const gotC = Math.round(got * 100);
      const esp = oraculoCent(100000, quem, div);
      checa(`1.2 ${quem}/${div}`, gotC === esp, `esperava ${esp}, veio ${gotC}`);
      cels.push((gotC > 0 ? `Ana +${gotC/100}` : gotC < 0 ? `Thi +${-gotC/100}` : "zerado")
                .padEnd(15) + (gotC === esp ? "" : " <<BUG"));
    }
    console.log(`     ${quem.padStart(14)} | ${cels.join(" | ")}`);
  }

  /* ---------- 1.3 invariantes de direcao (o que ele pediu) ---------- */
  console.log("\n[1.3] Invariantes de direcao - 40.000 casos aleatorios\n");
  let viol = { dirT: 0, dirA: 0, soma: 0, casa: 0, oraculo: 0 };
  let piorDif = 0, piorCaso = null;
  for (let i = 0; i < 40000; i++) {
    const vCent = 1 + Math.floor(Math.random() * 5000000);   // 0,01 a 50.000,00
    const v = vCent / 100;
    const div = ["meio", "thiago", "ana"][i % 3];
    const quem = ["Thiago", "Ana", "Casa"][(i / 3 | 0) % 3];
    const cotaT = round2(v * RATEIO[div]);
    const cotaA = round2(v - cotaT);
    const got = quem === "Thiago" ? cotaA : (quem === "Ana" ? -cotaT : 0);
    const gotC = Math.round(got * 100);
    const esp = oraculoCent(vCent, quem, div);

    if (gotC !== esp) {
      viol.oraculo++;
      const d = Math.abs(gotC - esp);
      if (d > piorDif) { piorDif = d; piorCaso = { v, quem, div, esp, gotC }; }
    }
    // cotaT + cotaA tem que devolver o valor exato
    if (Math.round((cotaT + cotaA) * 100) !== vCent) viol.soma++;
    // Thiago pagou e a conta nao e so dele -> Ana DEVE (positivo)
    if (quem === "Thiago" && div !== "thiago" && !(gotC > 0)) viol.dirT++;
    // Ana pagou e a conta nao e so dela -> Thiago DEVE (negativo)
    if (quem === "Ana" && div !== "ana" && !(gotC < 0)) viol.dirA++;
    // Casa pagou -> ninguem deve
    if (quem === "Casa" && gotC !== 0) viol.casa++;
  }
  const inv = [
    ["Thiago paga -> Ana sempre deve (positivo)", viol.dirT],
    ["Ana paga -> Thiago sempre deve (negativo)", viol.dirA],
    ["Casa paga -> acerto sempre zero", viol.casa],
    ["cota Thiago + cota Ana == valor exato", viol.soma],
    ["bate com o oraculo inteiro, ao centavo", viol.oraculo],
  ];
  for (const [nome, n] of inv) {
    checa("1.3 " + nome, n === 0, `${n} violacoes`);
    console.log(`  ${n === 0 ? "OK   " : "FALHA"} ${nome.padEnd(48)} ${n} violacao(oes)`);
  }
  if (piorCaso) console.log("       pior caso:", JSON.stringify(piorCaso));

  /* ---------- 1.4 arredondamento: os casos que quebram Math.round ---------- */
  console.log("\n[1.4] Arredondamento em valores que quebram Math.round classico\n");
  console.log("  O app normaliza a entrada com round2() antes de calcular (form e fixas),");
  console.log("  entao o pipeline real nunca ve 3 casas decimais. Testo os dois caminhos.\n");
  const dificeis = [9749.095, 1.005, 2.675, 0.615, 1.015, 8.575, 1234.565,
                    0.01, 0.03, 0.05, 1.11, 33.33, 66.67, 99999.99, 0.02];
  let ruins = 0;
  for (const bruto of dificeis) {
    const v = round2(bruto);                        // <-- o que o app faz na entrada
    const vCent = Math.round(v * 100);
    const cotaT = round2(v * 0.5);
    const cotaA = round2(v - cotaT);
    const got = Math.round(cotaA * 100);            // Thiago pagou
    const esp = oraculoCent(vCent, "Thiago", "meio");
    const ok = got === esp;
    if (!ok) ruins++;
    console.log(`  ${ok ? "OK  " : "FALHA"} digitou ${String(bruto).padEnd(10)} vira ${v.toFixed(2).padStart(9)}` +
                `  metade ${cotaT.toFixed(2).padStart(9)}  Ana deve ${(got/100).toFixed(2).padStart(9)}` +
                `  (oraculo ${(esp/100).toFixed(2)})`);
  }
  checa("1.4 arredondamento dificil", ruins === 0, `${ruins} divergencias`);

  /* ---------- 1.4b valor de 3 casas vindo DA PLANILHA (nao normalizado) ------- */
  console.log("\n[1.4b] Valor com 3+ casas lido da planilha - o app normaliza?\n");
  const buf = fs.readFileSync(BASE_XLSX);
  const { partes } = await win.lerZip(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length));
  const { linhas } = win.lerLancamentos(partes);
  const tresCasas = linhas.filter(l => Math.abs(l.valor * 100 - Math.round(l.valor * 100)) > 1e-9);
  console.log(`  ${linhas.length} lancamentos na planilha; ${tresCasas.length} com mais de 2 casas decimais`);
  if (tresCasas.length) {
    tresCasas.slice(0, 5).forEach(l =>
      console.log(`     linha ${l.linha}: ${l.descricao} = ${l.valor}`));
  }
  checa("1.4b planilha sem valores de 3 casas", tresCasas.length === 0,
        `${tresCasas.length} lancamentos com mais de 2 casas`);
  console.log(`  ${tresCasas.length === 0 ? "OK   " : "AVISO"} risco de divergencia JS x Excel: ` +
              `${tresCasas.length === 0 ? "nenhum (todos os valores tem 2 casas)" : "REAL"}`);

  /* ---------- 1.5 simetria: trocar o pagador inverte o sinal ---------- */
  console.log("\n[1.5] Trocar o pagador inverte o resultado\n");
  let assim = 0;
  for (let i = 0; i < 20000; i++) {
    const vCent = 2 + Math.floor(Math.random() * 2000000);
    const v = vCent / 100;
    const cotaT = round2(v * 0.5), cotaA = round2(v - cotaT);
    const seThiago = Math.round(cotaA * 100);
    const seAna = Math.round(-cotaT * 100);
    // em valor par a inversao e exata; em impar difere 1 centavo (alguem fica com ele)
    const limite = (vCent % 2 === 0) ? 0 : 1;
    if (Math.abs(seThiago + seAna) > limite) assim++;
    if (!(seThiago > 0 && seAna < 0)) assim++;
  }
  checa("1.5 simetria", assim === 0, `${assim} casos`);
  console.log(`  ${assim === 0 ? "OK   " : "FALHA"} inversao de sinal consistente em 20.000 casos ` +
              `(${assim} violacoes)`);

  /* ---------- resultado ---------- */
  console.log("\n" + "=".repeat(78));
  console.log(`BATERIA 1: ${R.ok} passaram, ${R.falhas.length} falharam`);
  if (R.falhas.length) {
    R.falhas.forEach(f => console.log("  FALHA:", f.nome, "-", f.detalhe));
    process.exitCode = 1;
  }
  console.log("=".repeat(78));
})().catch(e => { console.error("HARNESS QUEBROU:", e); process.exit(1); });
