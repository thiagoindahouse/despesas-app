/* ============================================================================
   BATERIA 2 - a interface. Abre a tela de fixas e o formulario no DOM real,
   mexe nos controles como o Thiago mexe, e confere o que aparece na tela.
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
const cent = s => {   // "R$ 1.234,56" -> 123456
  const m = String(s).replace(/[^\d,\-]/g, "").replace(",", ".");
  return Math.round(parseFloat(m) * 100);
};

(async () => {
  const { win } = await abrirApp();
  const doc = win.document;
  const $ = id => doc.getElementById(id);
  const V = expr => win.eval(expr);
  const disparar = (el, tipo) => el.dispatchEvent(new win.Event(tipo, { bubbles: true }));

  console.log("=".repeat(78));
  console.log("BATERIA 2 - INTERFACE (DOM real)");
  console.log("=".repeat(78));

  const buf = fs.readFileSync(BASE_XLSX);
  const fonte = fonteFake(buf, "base-teste.xlsx");
  await win.carregar(fonte);
  console.log(`\n  planilha carregada: ${V("ESTADO").dados.length} lancamentos, ` +
              `${V("ESTADO").recorrentes.length} recorrentes`);

  /* ---------- 2.1 a coluna de porcentagem sumiu mesmo? ---------- */
  console.log("\n[2.1] A coluna '% Thiago' sumiu da interface?\n");
  win.abrirFixas();
  const cabecalho = [...$("overlayFixas").querySelectorAll("th")].map(t => t.textContent.trim());
  console.log("  colunas da tela de fixas:", JSON.stringify(cabecalho));
  checa("2.1 sem % na tela", !cabecalho.some(c => /%/.test(c)),
        "ainda existe uma coluna com % na tela de fixas");
  checa("2.1 tem Divisao", cabecalho.some(c => /Divis/i.test(c)), "nao achei a coluna Divisao");
  checa("2.1 tem Resultado", cabecalho.some(c => /Resultado/i.test(c)), "nao achei a coluna Resultado");
  const opcoes = [...$("fxCorpo").querySelectorAll("select.div")[0].options].map(o => o.textContent);
  console.log("  opcoes de divisao:", JSON.stringify(opcoes));
  checa("2.1 opcoes em portugues claro",
        opcoes.join("|") === "Meio a meio|Conta do Thiago|Conta da Ana",
        "rotulos inesperados: " + opcoes.join("|"));

  /* ---------- 2.2 mexer no pagador muda o Resultado na hora ---------- */
  console.log("\n[2.2] Trocar o pagador inverte o Resultado na tela\n");
  const linhas = [...$("fxCorpo").querySelectorAll("tr[data-i]")];
  console.log(`  ${linhas.length} fixas montadas para ${$("fxMes").value}`);
  // deixa so a primeira marcada, com valor redondo, para isolar
  linhas.forEach((tr, i) => {
    const inc = tr.querySelector(".inc");
    if (inc.checked !== (i === 0)) { inc.checked = (i === 0); }
  });
  const tr0 = linhas[0];
  tr0.querySelector(".val").value = "1000.00";
  tr0.querySelector(".div").value = "meio";
  const nome0 = tr0.querySelector("td:nth-child(2)").textContent.trim().split("\n")[0];

  for (const [quem, esperado, rotulo] of [
      ["Thiago", 50000, "Ana deve 500,00"],
      ["Ana", -50000, "Thiago deve 500,00"],
      ["Casa", 0, "sem acerto"]]) {
    tr0.querySelector(".quem").value = quem;
    disparar(tr0.querySelector(".quem"), "input");
    const td = tr0.querySelector(".resultado").textContent.trim();
    const rodape = $("fxResumo").textContent.trim();
    let got;
    if (/sem acerto/.test(td)) got = 0;
    else got = (/^Ana/.test(td) ? 1 : -1) * cent(td);
    const ok = checa(`2.2 ${quem}`, got === esperado,
                     `linha diz "${td.replace(/\s+/g," ")}", esperava ${rotulo}`);
    console.log(`  ${ok ? "OK  " : "FALHA"} ${quem.padEnd(7)} paga R$ 1.000,00 -> ` +
                `linha: "${td.replace(/\s+/g, " ")}"`);
    console.log(`         rodape: ${rodape}`);
    // o rodape tem que concordar com a linha
    if (esperado !== 0) {
      const rOk = new RegExp(esperado > 0 ? "Ana deve" : "Thiago deve").test(rodape) &&
                  cent(rodape.split("·").pop()) === Math.abs(esperado);
      checa(`2.2 rodape ${quem}`, rOk, `rodape "${rodape}" nao bate com a linha`);
    }
  }

  /* ---------- 2.3 mexer na divisao ---------- */
  console.log("\n[2.3] Trocar a divisao muda o Resultado na tela\n");
  tr0.querySelector(".quem").value = "Thiago";
  for (const [div, esperado] of [["meio", 50000], ["thiago", 0], ["ana", 100000]]) {
    tr0.querySelector(".div").value = div;
    disparar(tr0.querySelector(".div"), "input");
    const td = tr0.querySelector(".resultado").textContent.trim();
    const got = /sem acerto/.test(td) ? 0 : (/^Ana/.test(td) ? 1 : -1) * cent(td);
    const rot = win.eval("DIVISOES").find(d => d.id === div).rotulo;
    const ok = checa(`2.3 ${div}`, got === esperado, `veio "${td.replace(/\s+/g," ")}", esperava ${esperado}`);
    console.log(`  ${ok ? "OK  " : "FALHA"} Thiago paga R$ 1.000,00, ${rot.padEnd(16)} -> ` +
                `"${td.replace(/\s+/g, " ")}"`);
  }

  /* ---------- 2.4 o cenario exato do print do Thiago ---------- */
  console.log("\n[2.4] Cenario real: setembro/2026 do print\n");
  const cenario = [
    ["Financiamento AP",   8413.96, "Thiago"],
    ["Cartão Casa Thiago", 5308.03, "Thiago"],
    ["Parcela Carro BMW",  3489.51, "Thiago"],
    ["Condomínio",         2700.00, "Thiago"],
    ["Faxina",             1100.00, "Thiago"],
    ["Cartão Porto",        748.27, "Thiago"],
    ["Gás",                 266.47, "Thiago"],
    ["Cartão Casa Ana",    6401.38, "Ana"],
    ["Luz",                 310.00, "Ana"],
  ];
  let somaT = 0, somaA = 0, imparesT = 0, imparesA = 0;
  cenario.forEach(([, v, q]) => {
    const impar = Math.round(v * 100) % 2 !== 0;
    if (q === "Thiago") { somaT += v; if (impar) imparesT++; }
    else { somaA += v; if (impar) imparesA++; }
  });
  const totalC = Math.round((somaT + somaA) * 100);
  const metadeC = Math.round(totalC / 2);
  const espAnaDeve = Math.round(somaT * 100) - metadeC;
  /* Arredondar linha a linha (que e o que a planilha ORIGINAL tambem faz, via
     ROUND na formula) desvia meio centavo por linha de valor impar. Com N linhas
     impares o desvio maximo e N/2 centavos. Isso nao e bug, e consequencia de
     arredondar cada linha em vez do total.                                     */
  const tolerancia = Math.ceil((imparesT + imparesA) / 2);
  console.log(`  Thiago pagou R$ ${somaT.toFixed(2)} | Ana pagou R$ ${somaA.toFixed(2)}`);
  console.log(`  total R$ ${(totalC/100).toFixed(2)} | metade de cada R$ ${(metadeC/100).toFixed(2)}`);
  console.log(`  conta a mao -> ANA DEVE R$ ${(espAnaDeve/100).toFixed(2)}`);
  console.log(`  ${imparesT + imparesA} linha(s) com centavo impar -> tolerancia ${tolerancia} centavo(s)\n`);

  // monta na tela: casa cada item do cenario com a linha CERTA pela descricao,
  // porque fixasDoMes() devolve na ordem da aba Recorrentes, nao na minha ordem
  const porDesc = new Map();
  linhas.forEach(tr => {
    const d = tr.querySelector("td:nth-child(2)").textContent.trim().split("\n")[0].trim();
    porDesc.set(d, tr);
  });
  console.log("  fixas oferecidas pelo app:", JSON.stringify([...porDesc.keys()]));
  const faltando = cenario.filter(([d]) => !porDesc.has(d)).map(([d]) => d);
  checa("2.4 todas as fixas do cenario existem", faltando.length === 0,
        "o app nao ofereceu: " + faltando.join(", "));
  const usadas = cenario.map(([d]) => porDesc.get(d)).filter(Boolean);
  linhas.forEach(tr => { tr.querySelector(".inc").checked = false; });
  cenario.forEach(([d, v, q]) => {
    const tr = porDesc.get(d);
    if (!tr) return;
    tr.querySelector(".inc").checked = true;
    tr.querySelector(".val").value = v.toFixed(2);
    tr.querySelector(".quem").value = q;
    tr.querySelector(".div").value = "meio";
  });
  disparar(usadas[0].querySelector(".val"), "input");

  cenario.forEach(([d, v, q]) => {
    const tr = porDesc.get(d);
    if (!tr) return;
    const td = tr.querySelector(".resultado").textContent.trim().replace(/\s+/g, " ");
    console.log(`     ${d.padEnd(20)} R$ ${v.toFixed(2).padStart(9)} (${q.padEnd(6)}) -> ${td}`);
  });
  const rodape = $("fxResumo").textContent.trim().replace(/\s+/g, " ");
  console.log(`\n  RODAPE: ${rodape}`);
  const gotFinal = /Ana deve/.test(rodape) ? cent(rodape.split("Ana deve").pop())
                 : -cent(rodape.split("Thiago deve").pop());
  const ok24 = checa("2.4 acerto do cenario", Math.abs(gotFinal - espAnaDeve) <= tolerancia,
        `tela diz ${gotFinal}, conta a mao da ${espAnaDeve}`);
  console.log(`  ${ok24 ? "OK  " : "FALHA"} tela: R$ ${(gotFinal/100).toFixed(2)} | ` +
              `mao: R$ ${(espAnaDeve/100).toFixed(2)} | diferenca ${Math.abs(gotFinal-espAnaDeve)} centavo(s)`);
  checa("2.4 direcao", gotFinal > 0, "a tela esta dizendo que o THIAGO deve, deveria ser a ANA");

  /* ---------- 2.5 grava de verdade e confere o que foi pro arquivo -------- */
  console.log("\n[2.5] Gravar no xlsx e reler - o que entrou e o que sai\n");
  const regs = win.lerFixasDaTela();
  checa("2.5 rateios gravados", regs.every(r => Math.abs(r.rateio - 0.5) < 1e-9),
        "algum rateio saiu diferente de 0.5: " + JSON.stringify(regs.map(r => r.rateio)));
  console.log(`  ${regs.length} registros, rateios: ${[...new Set(regs.map(r => r.rateio))].join(", ")}`);

  const antes = V("ESTADO").dados.length;
  await win.gravarLancamentos(fonte, regs, V("ESTADO").versao);
  await win.carregar(fonte);
  const novos = V("ESTADO").dados.slice(-cenario.length);
  console.log(`  gravou: ${antes} -> ${V("ESTADO").dados.length} lancamentos`);

  let somaLidos = 0, erros25 = 0;
  const espPorDesc = new Map(cenario.map(([d, v, q]) => [d, { v, q }]));
  novos.forEach(l => {
    somaLidos += l.anaDeve;
    const esp = espPorDesc.get(l.descricao);
    if (!esp) { erros25++; console.log(`     FALHA descricao inesperada: ${l.descricao}`); return; }
    const bate = Math.abs(l.valor - esp.v) < 0.005 && l.quem === esp.q &&
                 Math.abs(l.rateio - 0.5) < 1e-9;
    if (!bate) {
      erros25++;
      console.log(`     FALHA ${l.descricao}: gravou valor=${l.valor} quem=${l.quem} rateio=${l.rateio}, ` +
                  `esperava valor=${esp.v} quem=${esp.q} rateio=0.5`);
    }
  });
  checa("2.5 todas as descricoes voltaram", novos.length === cenario.length,
        `gravou ${novos.length}, esperava ${cenario.length}`);
  checa("2.5 round-trip fiel", erros25 === 0, `${erros25} linhas voltaram diferentes`);
  if (erros25 === 0) console.log(`  OK   as ${novos.length} linhas voltaram identicas (valor, pagador e rateio)`);
  const lidoC = Math.round(somaLidos * 100);
  const ok25 = checa("2.5 acerto apos reler", Math.abs(lidoC - espAnaDeve) <= tolerancia,
        `depois de gravar e reler deu ${lidoC}, esperava ${espAnaDeve}`);
  console.log(`  ${ok25 ? "OK  " : "FALHA"} acerto relido do arquivo: R$ ${(lidoC/100).toFixed(2)} ` +
              `(esperado R$ ${(espAnaDeve/100).toFixed(2)})`);

  fs.writeFileSync(require("path").join(__dirname, "saida-b2.xlsx"), fonte.buffer);
  console.log("  arquivo gravado em saida-b2.xlsx para a bateria do Excel");

  /* ---------- 2.6 formulario manual ---------- */
  console.log("\n[2.6] Formulario de lancamento avulso\n");
  win.abrirForm();
  const cab = [...$("overlay").querySelectorAll("label")].map(l => l.textContent.trim());
  checa("2.6 form sem %", !cab.some(c => /%/.test(c)), "o form ainda tem um campo com %");
  console.log("  campos do form:", JSON.stringify(cab.filter(c => c.length < 30)));
  $("fValor").value = "1000.00";
  $("fDesc").value = "Teste";
  for (const [quem, esperado] of [["Thiago", "Ana deve"], ["Ana", "Thiago deve"], ["Casa", "Ninguém deve nada"]]) {
    $("fQuem").value = quem;
    const selDiv = $("overlay").querySelector("select[id*='ivis'],#fDivisao,#fDiv");
    if (selDiv) selDiv.value = "meio";
    win.atualizarPrevia();
    const previa = $("previaRateio").textContent.trim().replace(/\s+/g, " ");
    const ok = checa(`2.6 previa ${quem}`,
                     new RegExp(esperado.replace(" ", "\\s+"), "i").test(previa) ||
                     (quem === "Casa" && /Ningu[eé]m deve nada/i.test(previa)),
                     `previa "${previa}" nao contem "${esperado}"`);
    console.log(`  ${ok ? "OK  " : "FALHA"} ${quem.padEnd(7)} -> ${previa}`);
  }

  console.log("\n" + "=".repeat(78));
  console.log(`BATERIA 2: ${R.ok} passaram, ${R.falhas.length} falharam`);
  R.falhas.forEach(f => console.log("  FALHA:", f.nome, "-", f.detalhe));
  console.log("=".repeat(78));
  if (R.falhas.length) process.exitCode = 1;
})().catch(e => { console.error("HARNESS QUEBROU:", e); process.exit(1); });



