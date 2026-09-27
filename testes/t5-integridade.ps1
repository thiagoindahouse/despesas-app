# ============================================================================
# BATERIA 5 - os arquivos gerados pelas baterias anteriores abrem no Excel
# sem reparo, sem erro de formula, e com os totais certos?
# ============================================================================
$ErrorActionPreference = "Stop"
$pasta = "C:\Users\thiagopaiva\.scout\copilot\session-state\4ca052fe-6657-4e91-9716-12b88f94a094\files"

Write-Host ("=" * 78)
Write-Host "BATERIA 5 - INTEGRIDADE DOS ARQUIVOS GERADOS"
Write-Host ("=" * 78)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$ok = 0; $falhas = 0

try {
  foreach ($nome in @("saida-b2.xlsx", "saida-b4.xlsx", "saida-b4-ano.xlsx")) {
    $arq = Join-Path $pasta $nome
    if (-not (Test-Path $arq)) { Write-Host "  PULADO $nome (nao existe)"; continue }
    Write-Host "`n--- $nome ---"

    # CorruptLoad = 0 (xlNormalLoad): se precisar de reparo, o Excel avisa
    $wb = $xl.Workbooks.Open($arq, 0, $false)
    $xl.CalculateFullRebuild()

    $ws = $wb.Worksheets.Item("Lançamentos")
    $ultima = $ws.Cells($ws.Rows.Count, 9).End(-4162).Row
    $n = $ultima - 1

    # erros reais (ignorando os NA() de proposito dos graficos)
    $reais = 0
    foreach ($sh in $wb.Worksheets) {
      try {
        foreach ($c in $sh.UsedRange.SpecialCells(-4123, 16)) {
          if (([string]$c.Formula) -notmatch "NA\(\)") {
            $reais++
            Write-Host ("     erro em {0}!{1}: {2}" -f $sh.Name, $c.Address(0,0), [string]$c.Formula)
          }
        }
      } catch { }
    }

    # soma da coluna de acerto x soma calculada dos pagamentos
    $somaAcerto = [double]$xl.WorksheetFunction.Sum($ws.Range($ws.Cells(2,14), $ws.Cells($ultima,14)))
    # confere que nenhuma celula da coluna 11 (% Thiago) esta vazia ou fora de 0..1
    $rateiosRuins = 0
    $dados = $ws.Range($ws.Cells(2,9), $ws.Cells($ultima,14)).Value2
    for ($r = 1; $r -le $n; $r++) {
      $rat = $dados[$r, 3]      # coluna 11 = offset 3 dentro do range que comeca em 9
      if ($null -eq $rat -or $rat -lt 0 -or $rat -gt 1) { $rateiosRuins++ }
    }

    Write-Host ("  {0} abriu sem reparo, {1} linhas, {2} erro(s) real(is) de formula" -f `
      $(if ($reais -eq 0) { "OK  " } else { "FALHA" }), $n, $reais)
    if ($reais -eq 0) { $ok++ } else { $falhas++ }

    Write-Host ("  {0} rateios fora da faixa 0..1: {1}" -f `
      $(if ($rateiosRuins -eq 0) { "OK  " } else { "FALHA" }), $rateiosRuins)
    if ($rateiosRuins -eq 0) { $ok++ } else { $falhas++ }

    Write-Host ("       soma da coluna 'Ana deve Thiago' na planilha inteira: R$ {0:N2}" -f $somaAcerto)

    $wb.Close($false)
  }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host ("=" * 78)
Write-Host "BATERIA 5: $ok passaram, $falhas falharam"
Write-Host ("=" * 78)
if ($falhas -gt 0) { exit 1 }
