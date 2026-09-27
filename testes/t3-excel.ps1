# ============================================================================
# BATERIA 3 - O Excel de verdade. Abre o arquivo que o app gravou, forca o
# recalculo e compara celula a celula com o que o JS calculou.
# Se o Excel discordar do app, a planilha mente para o Thiago.
# ============================================================================
$ErrorActionPreference = "Stop"
$pasta = "C:\Users\thiagopaiva\.scout\copilot\session-state\4ca052fe-6657-4e91-9716-12b88f94a094\files"
$arq = Join-Path $pasta "saida-b2.xlsx"

Write-Host ("=" * 78)
Write-Host "BATERIA 3 - EXCEL REAL"
Write-Host ("=" * 78)

if (-not (Test-Path $arq)) { throw "rode a bateria 2 antes: $arq nao existe" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$falhas = 0
$ok = 0

try {
  $wb = $xl.Workbooks.Open($arq, 0, $false)
  Write-Host "`n  abriu sem reparo. Abas: $($wb.Worksheets.Count)"
  $xl.CalculateFullRebuild()
  Write-Host "  recalculo completo forcado"

  $ws = $wb.Worksheets.Item("Lançamentos")
  $ultima = $ws.Cells($ws.Rows.Count, 9).End(-4162).Row
  Write-Host "  aba Lancamentos, ultima linha com valor: $ultima"

  # ---- 3.1 erros de formula em TODA a pasta ----
  Write-Host "`n[3.1] Erros de formula em toda a pasta de trabalho`n"
  Write-Host "  (as linhas de apoio dos graficos usam NA() de proposito, para abrir"
  Write-Host "   buraco no grafico nos meses sem dado; esses nao contam como erro)`n"
  $totalErros = 0
  foreach ($sh in $wb.Worksheets) {
    $usada = $sh.UsedRange
    $reais = 0; $intencionais = 0
    try {
      $cels = $usada.SpecialCells(-4123, 16)
      foreach ($c in $cels) {
        $f = [string]$c.Formula
        if ($f -match "NA\(\)") { $intencionais++ } else {
          $reais++
          Write-Host ("       erro real em {0}!{1}: {2} = {3}" -f `
            $sh.Name, $c.Address(0,0), $f, [string]$c.Text)
        }
      }
    } catch { }
    $marca = if ($reais -eq 0) { "OK   " } else { "FALHA" }
    Write-Host ("  {0} {1,-18} {2} erro(s) real(is), {3} NA() intencional(is)" -f `
      $marca, $sh.Name, $reais, $intencionais)
    $totalErros += $reais
  }
  if ($totalErros -eq 0) { $ok++ } else { $falhas++ }

  # ---- 3.2 as 9 linhas novas: Excel x JS ----
  Write-Host "`n[3.2] As 9 linhas de setembro: o que o Excel calcula`n"
  $esperado = @{
    "Financiamento AP"   = @(8413.96, "Thiago", 4206.98)
    "Cartão Casa Thiago" = @(5308.03, "Thiago", 2654.01)
    "Parcela Carro BMW"  = @(3489.51, "Thiago", 1744.75)
    "Condomínio"         = @(2700.00, "Thiago", 1350.00)
    "Faxina"             = @(1100.00, "Thiago", 550.00)
    "Cartão Porto"       = @(748.27,  "Thiago", 374.13)
    "Gás"                = @(266.47,  "Thiago", 133.23)
    "Cartão Casa Ana"    = @(6401.38, "Ana",    -3200.69)
    "Luz"                = @(310.00,  "Ana",    -155.00)
  }
  $somaExcel = 0.0
  $inicio = $ultima - 8
  Write-Host ("  {0,-20} {1,10} {2,-7} {3,11} {4,11}" -f "Descricao","Valor","Pagou","Excel","Esperado")
  Write-Host ("  " + ("-" * 74))
  for ($r = $inicio; $r -le $ultima; $r++) {
    $desc = [string]$ws.Cells($r, 7).Value2
    $val  = [double]$ws.Cells($r, 9).Value2
    $quem = [string]$ws.Cells($r, 10).Value2
    $rat  = [double]$ws.Cells($r, 11).Value2
    $ad   = [double]$ws.Cells($r, 14).Value2      # coluna "Ana deve Thiago"
    $somaExcel += $ad
    $e = $esperado[$desc]
    if ($null -eq $e) {
      Write-Host ("  FALHA linha {0}: descricao inesperada '{1}'" -f $r, $desc); $falhas++; continue
    }
    $difV = [math]::Abs($val - $e[0])
    $difA = [math]::Abs($ad - $e[2])
    $bate = ($difV -lt 0.005) -and ($quem -eq $e[1]) -and ($difA -lt 0.005) -and ([math]::Abs($rat - 0.5) -lt 1e-9)
    if ($bate) { $ok++ } else { $falhas++ }
    Write-Host ("  {0} {1,-20} {2,10:N2} {3,-7} {4,11:N2} {5,11:N2}" -f `
      $(if ($bate) { "OK  " } else { "FALHA" }), $desc, $val, $quem, $ad, $e[2])
  }

  # ---- 3.3 o acerto total segundo o Excel ----
  Write-Host "`n[3.3] Acerto de setembro segundo o EXCEL`n"
  Write-Host ("  soma da coluna 'Ana deve Thiago': R$ {0:N2}" -f $somaExcel)
  $espTotal = 7657.41
  $bate33 = [math]::Abs($somaExcel - $espTotal) -lt 0.005
  if ($bate33) { $ok++ } else { $falhas++ }
  Write-Host ("  {0} Excel: R$ {1:N2} | app (JS): R$ {2:N2}" -f `
    $(if ($bate33) { "OK  " } else { "FALHA" }), $somaExcel, $espTotal)
  if ($somaExcel -gt 0) {
    Write-Host ("  DIRECAO: positivo -> ANA DEVE ao Thiago  <- e o que o Thiago pediu")
    $ok++
  } else {
    Write-Host ("  FALHA DIRECAO: negativo -> Thiago deve a Ana  <- INVERTIDO")
    $falhas++
  }

  # ---- 3.4 as formulas gravadas sao as certas? ----
  Write-Host "`n[3.4] Texto das formulas gravadas (nao pode ter #REF!)`n"
  foreach ($c in 12, 13, 14) {
    $f = [string]$ws.Cells($ultima, $c).Formula
    $nome = [string]$ws.Cells(1, $c).Value2
    $temRef = $f -match "#REF"
    if ($temRef) { $falhas++ } else { $ok++ }
    Write-Host ("  {0} {1,-18} {2}" -f $(if ($temRef) { "FALHA" } else { "OK  " }), $nome, $f)
  }

  # ---- 3.5 o Painel continua consistente ----
  Write-Host "`n[3.5] Painel: recalcula sem erro e reage aos dados novos`n"
  $pn = $wb.Worksheets.Item("Painel")
  $nErrPainel = 0
  try { $nErrPainel = 0; foreach ($c in $pn.UsedRange.SpecialCells(-4123,16)) { if (([string]$c.Formula) -notmatch "NA\(\)") { $nErrPainel++ } } } catch { $nErrPainel = 0 }
  if ($nErrPainel -eq 0) { $ok++ } else { $falhas++ }
  Write-Host ("  {0} Painel com {1} erro(s) de formula" -f $(if ($nErrPainel -eq 0) { "OK  " } else { "FALHA" }), $nErrPainel)

  $wb.Close($false)
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host ("=" * 78)
Write-Host "BATERIA 3: $ok passaram, $falhas falharam"
Write-Host ("=" * 78)
if ($falhas -gt 0) { exit 1 }

