# Roda a bateria inteira do zero e consolida o resultado.
$ErrorActionPreference = "Continue"
$pasta = "C:\Users\thiagopaiva\.scout\copilot\session-state\4ca052fe-6657-4e91-9716-12b88f94a094\files"
Set-Location $pasta

# limpa saidas de rodadas anteriores para nao mascarar falha
Remove-Item "$pasta\saida-b*.xlsx" -Force -ErrorAction SilentlyContinue
Copy-Item "C:\Users\thiagopaiva\OneDrive\Despesas Ana e Thiago\Despesas Ana e Thiago.xlsx" `
          "$pasta\base-teste.xlsx" -Force

$baterias = @(
  @{ n = "1. Regra central";            cmd = "node";       arg = "t1-regra.js" },
  @{ n = "2. Interface (DOM)";          cmd = "node";       arg = "t2-tela.js" },
  @{ n = "3. Excel real";               cmd = "powershell"; arg = "-ExecutionPolicy Bypass -File t3-excel.ps1" },
  @{ n = "4. Robustez e extremos";      cmd = "node";       arg = "t4-robustez.js" },
  @{ n = "5. Integridade dos arquivos"; cmd = "powershell"; arg = "-ExecutionPolicy Bypass -File t5-integridade.ps1" },
  @{ n = "6. Reconciliacao historica";  cmd = "node";       arg = "t6-historico.js" },
  @{ n = "7. Ciclo completo do mes";    cmd = "node";       arg = "t7-ciclo.js" }
)

$res = @()
foreach ($b in $baterias) {
  Write-Host ("`n{0}`n  RODANDO: {1}`n{0}" -f ("-" * 78), $b.n)
  $saida = & $b.cmd $b.arg.Split(" ") 2>&1 | Out-String
  $code = $LASTEXITCODE
  $linha = ($saida -split "`n" | Where-Object { $_ -match "^BATERIA \d+: " } | Select-Object -Last 1)
  if ($linha -match "(\d+) passaram, (\d+) falharam") {
    $p = [int]$Matches[1]; $f = [int]$Matches[2]
  } else { $p = 0; $f = -1 }
  $res += [pscustomobject]@{ Bateria = $b.n; Passaram = $p; Falharam = $f; Codigo = $code }
  Write-Host ("  -> {0}" -f $linha.Trim())
  if ($f -ne 0) {
    Write-Host "  DETALHE DAS FALHAS:"
    ($saida -split "`n" | Where-Object { $_ -match "FALHA" } | Select-Object -First 15) |
      ForEach-Object { Write-Host "    $_" }
  }
}

Write-Host ""
Write-Host ("=" * 78)
Write-Host "RESULTADO CONSOLIDADO"
Write-Host ("=" * 78)
$res | Format-Table -AutoSize
$totalP = ($res | Measure-Object Passaram -Sum).Sum
$totalF = ($res | Measure-Object Falharam -Sum).Sum
Write-Host ("TOTAL: {0} verificacoes passaram, {1} falharam" -f $totalP, $totalF)
if ($totalF -eq 0) {
  Write-Host "ZERO FALHAS."
} else {
  Write-Host "AINDA HA FALHA - NAO ENTREGAR."
  exit 1
}
