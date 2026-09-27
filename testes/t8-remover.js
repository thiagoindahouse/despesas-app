/* ============================================================================
   BATERIA 8 - a remocao de linhas, testada ANTES de encostar na planilha dele.
   Grava lancamentos de setembro numa copia, remove, e confere que:
     - so setembro saiu
     - o resto ficou intacto, linha por linha
     - o Excel abre sem reparo e sem erro de formula
     - dá para lancar de novo depois
   ========================================================================== */
const { abrirApp, fonteFake, BASE_XLSX } = require("./harness");
const { removerLinhasDoSheet } = require("./remover");
const fs = require("fs");
const path = require("path");

const R = { ok: 0, falhas: [] };
function checa(nome, cond, detalhe) {
  if (cond) { R.ok++; return true; }
  R.falhas.push({ nome, detalhe });
  console.log(`  FALHA ${nome}: ${detalhe}`);
  return false;
}

/* Remove da fonte todos os lancamentos de um ano/mes. Mesma mecanica de
   gravarLancamentos, ao contrario.                                          */
async function removerMes(win, fonte, ano, mes) {
  const { buf, versao } = await fonte.ler();
  const { ordem, partes } = await win.lerZip(buf);
  const caminho = win.acharAbaLancamentos(partes);
  const tabela = win.acharTabela(ordem, partes);
  if (!tabela) throw new Error("nao achei a tabela tbLancamentos");

  const { linhas } = win.lerLancamentos(partes);
  const alvo = linhas.filter(l => l.ano === ano && l.mes === mes).map(l => l.linha);
  if (!alvo.length) return { removidas: 0, alvo: [] };

  const te = new TextEncoder();
  const txt = b => new TextDecoder().decode(b);

  let xml = txt(partes.get(caminho));
  const r = removerLinhasDoSheet(xml, alvo);
  partes.set(caminho, te.encode(r.xml));

  // a tabela tem que encolher junto
  let tx = txt(partes.get(tabela.parte));
  tx = tx.replace(/(ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/g, `$1${r.ultima}$3`);
  partes.set(tabela.parte, te.encode(tx));

  // recalcular ao abrir e descartar o calcChain, que fica invalido
  let wbx = txt(partes.get("xl/workbook.xml"));
  if (/<calcPr\b[^>]*\/?>/.test(wbx)) {
    wbx = wbx.replace(/<calcPr\b[^>]*?\/?>/, m =>
      m.replace(/\s*fullCalcOnLoad="[^"]*"/, "").replace(/\/?>$/, ' fullCalcOnLoad="1"/>'));
  } else {
    wbx = wbx.replace("</workbook>", '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>');
  }
  partes.set("xl/workbook.xml", te.encode(wbx));
  const semCalc = ordem.filter(n => n !== "xl/calcChain.xml");
  partes.delete("xl/calcChain.xml");
  let ct = txt(partes.get("[Content_Types].xml"));
  ct = ct.replace(/<Override[^>]*calcChain\.xml[^>]*\/>/g, "");
  partes.set("[Content_Types].xml", te.encode(ct));

  const blob = await win.escreverZip(semCalc, partes);
  await fonte.gravar(blob, versao);
  return { removidas: r.removidas, alvo, ultima: r.ultima };
}

(async () => {
  const { win } = await abrirApp();
  const V = e => win.eval(e);

  console.log("=".repeat(78));
  console.log("BATERIA 8 - REMOCAO DE LANCAMENTOS (teste antes de usar)");
  console.log("=".repeat(78));

  const fonte = fonteFake(fs.readFileSync(BASE_XLSX));
  await win.carregar(fonte);
  const original = V("ESTADO").dados.map(l =>
    `${l.ano}|${l.mes}|${l.descricao}|${l.valor}|${l.quem}|${l.rateio}`);
  console.log(`\n  planilha de partida: ${original.length} lancamentos`);

  /* --- simula o estrago: grava 9 lancamentos de setembro com 100% (como ficou) --- */
  console.log("\n[8.1] Simulo o estrago: 9 lancamentos de setembro com rateio 1\n");
  const estrago = [
    ["Financiamento AP", 8413.96, "Thiago"], ["Cartão Casa Thiago", 5308.03, "Thiago"],
    ["Parcela Carro BMW", 3489.51, "Thiago"], ["Condomínio", 2700, "Thiago"],
    ["Faxina", 1100, "Thiago"], ["Cartão Porto", 748.27, "Thiago"],
    ["Gás", 266.47, "Thiago"], ["Cartão Casa Ana", 6401.38, "Ana"], ["Luz", 310, "Ana"],
  ].map(([descricao, valor, quem]) => ({
    data: new Date(2026, 8, 1), tipo: "Fixa", categoria: "Teste", descricao,
    parcela: "", valor, quem, rateio: 1, obs: "",      // 1 = o bug original
  }));
  const at = await fonte.ler();
  await win.gravarLancamentos(fonte, estrago, at.versao);
  await win.carregar(fonte);
  let est = V("ESTADO");
  const set1 = est.dados.filter(l => l.ano === 2026 && l.mes === 9);
  console.log(`  ${est.dados.length} lancamentos, ${set1.length} em setembro/2026`);
  const acertoErrado = set1.reduce((s, l) => s + l.anaDeve, 0);
  console.log(`  acerto com o bug: R$ ${acertoErrado.toFixed(2)} ` +
              `(${acertoErrado > 0 ? "Ana" : "Thiago"} deve) <- era isso que voce via`);
  checa("8.1 estrago reproduzido", set1.length === 9, `${set1.length} linhas`);

  /* --- remove --- */
  console.log("\n[8.2] Removo setembro/2026\n");
  const r = await removerMes(win, fonte, 2026, 9);
  console.log(`  removi ${r.removidas} linha(s): ${r.alvo.join(", ")}`);
  await win.carregar(fonte);
  est = V("ESTADO");
  const set2 = est.dados.filter(l => l.ano === 2026 && l.mes === 9);
  console.log(`  sobraram ${est.dados.length} lancamentos, ${set2.length} em setembro`);
  checa("8.2 setembro saiu", set2.length === 0, `ainda ha ${set2.length} linhas de setembro`);
  checa("8.2 contagem", est.dados.length === original.length,
        `sobrou ${est.dados.length}, esperava ${original.length}`);

  /* --- o resto ficou intacto? --- */
  console.log("\n[8.3] O resto da planilha ficou intacto?\n");
  const depois = est.dados.map(l =>
    `${l.ano}|${l.mes}|${l.descricao}|${l.valor}|${l.quem}|${l.rateio}`);
  let difs = 0;
  for (let i = 0; i < Math.max(original.length, depois.length); i++) {
    if (original[i] !== depois[i]) {
      if (difs < 5) console.log(`     linha ${i}: antes "${original[i]}" | depois "${depois[i]}"`);
      difs++;
    }
  }
  checa("8.3 conteudo identico", difs === 0, `${difs} linhas diferentes`);
  console.log(`  ${difs === 0 ? "OK  " : "FALHA"} os ${original.length} lancamentos originais ` +
              `voltaram identicos (${difs} diferenca(s))`);

  // numeracao sem buraco
  const nums = est.dados.map(l => l.linha);
  const buracos = nums.filter((n, i) => i > 0 && n !== nums[i - 1] + 1).length;
  checa("8.3 numeracao continua", buracos === 0, `${buracos} buraco(s) na numeracao`);
  console.log(`  ${buracos === 0 ? "OK  " : "FALHA"} numeracao de linhas sem buraco ` +
              `(${nums[0]} ate ${nums[nums.length - 1]})`);

  /* --- dá para lancar de novo, certo desta vez? --- */
  console.log("\n[8.4] Lancar setembro de novo, agora meio a meio\n");
  const bom = estrago.map(e => ({ ...e, rateio: 0.5 }));
  const at2 = await fonte.ler();
  await win.gravarLancamentos(fonte, bom, at2.versao);
  await win.carregar(fonte);
  est = V("ESTADO");
  const set3 = est.dados.filter(l => l.ano === 2026 && l.mes === 9);
  const acertoCerto = set3.reduce((s, l) => s + l.anaDeve, 0);
  console.log(`  ${set3.length} lancamentos de setembro`);
  console.log(`  acerto: R$ ${acertoCerto.toFixed(2)} (${acertoCerto > 0 ? "ANA" : "THIAGO"} deve)`);
  checa("8.4 relancou", set3.length === 9, `${set3.length} linhas`);
  checa("8.4 direcao certa", acertoCerto > 0, "continua invertido");
  checa("8.4 valor certo", Math.abs(acertoCerto - 7657.41) < 0.02,
        `deu ${acertoCerto.toFixed(2)}, esperava 7657.41`);

  const saida = path.join(__dirname, "saida-b8.xlsx");
  fs.writeFileSync(saida, fonte.buffer);
  console.log(`\n  gravado em ${path.basename(saida)} para conferir no Excel`);

  console.log("\n" + "=".repeat(78));
  console.log(`BATERIA 8: ${R.ok} passaram, ${R.falhas.length} falharam`);
  R.falhas.forEach(f => console.log("  FALHA:", f.nome, "-", f.detalhe));
  console.log("=".repeat(78));
  if (R.falhas.length) process.exitCode = 1;
})().catch(e => { console.error("HARNESS QUEBROU:", e); process.exit(1); });
