# Gerador do Memorial de Calculos — Midway 13864630
# Simone Marques Martins Sepeda
# Formato: Word .docx pronto para juntada ao processo

$docPath = "C:\Users\CONCEIÇÃO DUARTE\teste\peticoes\memorial-calculos-simone-sepeda-midway-2026-06-17.docx"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc  = $word.Documents.Add()
$sel  = $word.Selection

# Margens da pagina (em pontos: 1 cm = 28.35 pt)
$doc.PageSetup.TopMargin    = 56.7   # 2 cm
$doc.PageSetup.BottomMargin = 56.7
$doc.PageSetup.LeftMargin   = 85.05  # 3 cm
$doc.PageSetup.RightMargin  = 56.7

# IDs numericos dos estilos built-in do Word (qualquer idioma)
$ST_NORMAL = -1
$ST_H1     = -2
$ST_H2     = -3
$ST_H3     = -4

function SetStyle($id) {
    try { $sel.Style = $doc.Styles.Item($id) } catch {}
}
function H1($t)  { SetStyle $ST_H1;     $sel.TypeText($t); $sel.TypeParagraph() }
function H2($t)  { SetStyle $ST_H2;     $sel.TypeText($t); $sel.TypeParagraph() }
function H3($t)  { SetStyle $ST_H3;     $sel.TypeText($t); $sel.TypeParagraph() }
function NL()    { SetStyle $ST_NORMAL; $sel.TypeParagraph() }
function Para($t){ SetStyle $ST_NORMAL; $sel.Font.Bold=$false; $sel.Font.Italic=$false; $sel.TypeText($t); $sel.TypeParagraph() }
function Bold($t){ SetStyle $ST_NORMAL; $sel.Font.Bold=$true;  $sel.TypeText($t); $sel.TypeParagraph(); $sel.Font.Bold=$false }
function Ital($t){ SetStyle $ST_NORMAL; $sel.Font.Italic=$true; $sel.TypeText($t); $sel.TypeParagraph(); $sel.Font.Italic=$false }

# Insere tabela 2 colunas com linhas passadas como array plano de pares
function Tabela2($col1, $col2, [string[]]$keys, [string[]]$vals) {
    $n   = $keys.Count
    $rng = $sel.Range
    $t   = $doc.Tables.Add($rng, $n+1, 2)
    $t.Borders.Enable = $true
    $t.Cell(1,1).Range.Text = $col1; $t.Cell(1,1).Range.Font.Bold = $true
    $t.Cell(1,2).Range.Text = $col2; $t.Cell(1,2).Range.Font.Bold = $true
    $t.Cell(1,1).Shading.BackgroundPatternColor = 14277081
    $t.Cell(1,2).Shading.BackgroundPatternColor = 14277081
    for ($i = 0; $i -lt $n; $i++) {
        $t.Cell($i+2, 1).Range.Text = [string]$keys[$i]
        $t.Cell($i+2, 2).Range.Text = [string]$vals[$i]
    }
    $sel.Start = $t.Range.End
    $sel.TypeParagraph()
}

# Insere tabela 3 colunas
function Tabela3($c1,$c2,$c3, [string[]]$a,[string[]]$b,[string[]]$c) {
    $n   = $a.Count
    $rng = $sel.Range
    $t   = $doc.Tables.Add($rng, $n+1, 3)
    $t.Borders.Enable = $true
    foreach ($col in 1..3) { $t.Cell(1,$col).Shading.BackgroundPatternColor = 14277081; $t.Cell(1,$col).Range.Font.Bold = $true }
    $t.Cell(1,1).Range.Text = $c1
    $t.Cell(1,2).Range.Text = $c2
    $t.Cell(1,3).Range.Text = $c3
    for ($i = 0; $i -lt $n; $i++) {
        $t.Cell($i+2,1).Range.Text = [string]$a[$i]
        $t.Cell($i+2,2).Range.Text = [string]$b[$i]
        $t.Cell($i+2,3).Range.Text = [string]$c[$i]
    }
    $sel.Start = $t.Range.End
    $sel.TypeParagraph()
}

# Insere tabela 8 colunas para amortizacao
function TabelaAmort() {
    $headers = @("Parc","Saldo(9,99%)","PMT(9,99%)","Juros(9,99%)","Saldo(5,10%)","PMT(5,10%)","Juros(5,10%)","Excesso")
    $parcs   = @("1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","TOTAL")
    $sd99    = @("R$ 5.163,35","R$ 5.082,58","R$ 4.993,74","R$ 4.896,02","R$ 4.788,54","R$ 4.670,33","R$ 4.540,30","R$ 4.397,29","R$ 4.239,99","R$ 4.066,97","R$ 3.876,68","R$ 3.667,36","R$ 3.437,14","R$ 3.183,92","R$ 2.905,41","R$ 2.599,07","R$ 2.262,12","R$ 1.891,52","R$ 1.483,89","R$ 1.035,54","R$ 542,40","—")
    $pmt99   = @("R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 596,59","R$ 12.528,40")
    $j99     = @("R$ 515,82","R$ 507,75","R$ 498,87","R$ 489,11","R$ 478,38","R$ 466,57","R$ 453,58","R$ 439,29","R$ 423,57","R$ 406,29","R$ 387,28","R$ 366,37","R$ 343,37","R$ 318,07","R$ 290,25","R$ 259,65","R$ 225,99","R$ 188,96","R$ 148,24","R$ 103,45","R$ 54,19","R$ 7.365,05")
    $sd51    = @("R$ 5.163,35","R$ 5.020,41","R$ 4.870,18","R$ 4.712,28","R$ 4.546,33","R$ 4.371,92","R$ 4.188,62","R$ 3.995,96","R$ 3.793,49","R$ 3.580,68","R$ 3.357,02","R$ 3.121,96","R$ 2.874,90","R$ 2.615,25","R$ 2.342,35","R$ 2.055,54","R$ 1.754,10","R$ 1.437,28","R$ 1.104,31","R$ 754,36","R$ 386,56","—")
    $pmt51   = @("R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 406,27","R$ 8.531,74")
    $j51     = @("R$ 263,33","R$ 256,04","R$ 248,38","R$ 240,33","R$ 231,86","R$ 222,97","R$ 213,62","R$ 203,79","R$ 193,47","R$ 182,61","R$ 171,21","R$ 159,22","R$ 146,62","R$ 133,38","R$ 119,46","R$ 104,83","R$ 89,46","R$ 73,30","R$ 56,32","R$ 38,47","R$ 19,71","R$ 3.368,39")
    $exc     = @("R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 190,32","R$ 3.996,66")
    $n  = $parcs.Count
    $rng = $sel.Range
    $t   = $doc.Tables.Add($rng, $n+1, 8)
    $t.Borders.Enable = $true
    for ($c = 1; $c -le 8; $c++) {
        $t.Cell(1,$c).Range.Text = $headers[$c-1]
        $t.Cell(1,$c).Range.Font.Bold = $true
        $t.Cell(1,$c).Shading.BackgroundPatternColor = 14277081
    }
    for ($i = 0; $i -lt $n; $i++) {
        $t.Cell($i+2,1).Range.Text = $parcs[$i]
        $t.Cell($i+2,2).Range.Text = $sd99[$i]
        $t.Cell($i+2,3).Range.Text = $pmt99[$i]
        $t.Cell($i+2,4).Range.Text = $j99[$i]
        $t.Cell($i+2,5).Range.Text = $sd51[$i]
        $t.Cell($i+2,6).Range.Text = $pmt51[$i]
        $t.Cell($i+2,7).Range.Text = $j51[$i]
        $t.Cell($i+2,8).Range.Text = $exc[$i]
    }
    $sel.Start = $t.Range.End
    $sel.TypeParagraph()
}

# ════════════════════════════════════════════════════════════════
# DOCUMENTO
# ════════════════════════════════════════════════════════════════

H1 "MEMORIAL DE CALCULOS"
H2 "Revisional — Midway Financial S.A. | Credito Pessoal no 13864630"
Para "Cliente: SIMONE MARQUES MARTINS SEPEDA | CPF: 324.863.672-68"
Para "Elaborado em: 17/06/2026   |   Advogada: Conceicao Maria Duarte Portilho"
Para "Fonte BACEN: Serie 25479 — api.bcb.gov.br/dados/serie/bcdata.sgs.25479"
NL

# 1
H2 "1. Dados do Contrato"
Tabela2 "Item" "Valor" `
  @("Principal liberado (especie)","IOF financiado","Principal total financiado","Taxa nominal declarada","CET declarado","PMT cobrado pelo banco","PMT PRICE calculado a 9,99%","DIVERGENCIA por parcela","Taxa efetiva real (TIR calculada)","Total PRICE a 9,99% (calculado)","Total contratual (636,64 x 21)","Excesso pela carencia/TIR real","Carencia","Parcelas") `
  @("R$ 5.000,00","R$ 163,35","R$ 5.163,35","9,99% a.m. / 213,50% a.a.","10,52% a.m. / 231,95% a.a.","R$ 636,64","R$ 596,59","R$ 40,05 a mais/mes","10,9352% a.m. / 247,40% a.a.","R$ 12.528,40","R$ 13.369,44","R$ 841,04","48 dias","21")
Ital "NOTA: O PMT de R$ 636,64 NAO corresponde a formula PRICE a 9,99% a.m. (que produziria R$ 596,59). A TIR real apurada e 10,9352% a.m. — 0,94 p.p. acima do declarado. Fundamento: CDC art. 6, III; Res. BACEN 3.517/2007."
NL

# 2
H2 "2. Referencia BACEN — Serie 25479"
Tabela2 "Dado" "Valor" `
  @("Serie","Mes de contratacao (jul/2024)","Equivalente anual","Fonte") `
  @("25479 — Credito pessoal nao consignado (outros)","5,10% a.m.","81,65% a.a.","api.bcb.gov.br/dados/serie/bcdata.sgs.25479")
NL

# 3
H2 "3. Abusividade — STJ Tema 27 / REsp 1.061.530/RS"
Tabela3 "Indice" "Mensal" "Anual" `
  @("Taxa declarada no contrato","Taxa efetiva real (TIR)","Taxa media BACEN — jul/2024","Excesso absoluto (TIR vs BACEN)","Excesso relativo (TIR vs BACEN)") `
  @("9,99%","10,9352%","5,10%","5,84 p.p./mes","114,4% acima") `
  @("213,50%","247,40%","81,65%","—","—")
Bold "-> Taxa efetiva real e mais que o DOBRO da taxa media de mercado."
Bold "-> Abusividade configurada — STJ REsp 1.061.530/RS (recurso repetitivo)."
NL

# 4
H2 "4. Tabela de Amortizacao Comparativa (PRICE)"
Para "Principal: R$ 5.163,35 | Contratada 9,99% a.m. x BACEN 5,10% a.m."
TabelaAmort
NL

# 5
H2 "5. Pagamentos Reais Efetuados (extrato Midway)"
Tabela2 "Parcela" "Valor pago" `
  @("01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","TOTAL PAGO") `
  @("R$ 626,62","R$ 669,60","R$ 662,79","R$ 630,61","R$ 636,64","R$ 660,58","R$ 634,62","R$ 636,64","R$ 636,64","R$ 636,64","R$ 636,64","R$ 634,62","R$ 636,64","R$ 632,61","R$ 628,61","R$ 628,61","R$ 626,62","R$ 630,61","R$ 626,62","R$ 584,35","R$ 578,98","R$ 13.276,29")
Ital "Contrato quitado em 16/04/2026. Saldo devedor: R$ 0,00."
NL

# 6
H2 "6. Quadro do Indbito — CC art. 876 + CDC art. 42"
Bold "Total efetivamente pago (extrato Midway): R$ 13.276,29"
NL

H3 "Cenario A — Taxa BACEN 5,10% a.m. (pedido principal)"
Tabela2 "Item" "Valor" `
  @("PMT revisado","Total devido a 5,10%","Juros pagos em excesso","INDBITO (A)","Restituicao em dobro — CDC 42 par. unico","(dobro somente se provada ma-fe — STJ Tema 929)") `
  @("R$ 406,27","R$ 8.531,74","R$ 3.996,66","R$ 4.744,55","R$ 9.489,10","")
NL

H3 "Cenario B — 7,00% a.m. com spread (pedido subsidiario)"
Tabela2 "Item" "Valor" `
  @("PMT revisado","Total devido a 7,00%","INDBITO (B)") `
  @("R$ 476,52","R$ 10.006,93","R$ 3.269,36")
NL

H3 "Cenario C — BACEN 5,10% sem IOF financiado (maximo)"
Tabela2 "Item" "Valor" `
  @("Base: R$ 5.000,00 | PMT revisado","Total devido","INDBITO (C)") `
  @("R$ 393,42","R$ 8.261,83","R$ 5.014,46")
NL

H3 "Resumo Consolidado"
Tabela3 "Cenario" "Indbito" "Base" `
  @("A — Pedido principal","B — Subsidiario conservador","C — Maximo (com IOF)","Em dobro — Cenario A (se ma-fe)") `
  @("R$ 4.744,55","R$ 3.269,36","R$ 5.014,46","R$ 9.489,10") `
  @("BACEN 5,10% s/ R$ 5.163,35","7,00% s/ R$ 5.163,35","BACEN 5,10% s/ R$ 5.000,00","Somente se ma-fe provada")
NL

# 7
H2 "7. Prescricao — CC art. 206, par. 5, I (5 anos por pagamento)"
Tabela3 "Marco" "Data" "Prescricao" `
  @("1a parcela paga","Ultima parcela paga","Data do calculo") `
  @("11/09/2024","16/04/2026","17/06/2026") `
  @("11/09/2029","16/04/2031","—")
Bold "-> Todos os pagamentos estao dentro do prazo prescricional. Restituicao integral cabivel."
NL

# 8
H2 "8. IOF Financiado — CDC art. 51, IV"
Tabela2 "Item" "Valor" `
  @("IOF embutido no principal","Juros pagos sobre o IOF (9,99% a.m.)") `
  @("R$ 163,35","R$ 233,00 (estimado proporcional)")
Ital "Cliente pagou juros de 9,99% a.m. sobre R$ 163,35 de tributo que nunca recebeu em especie. Pratica abusiva — CDC art. 51, IV."
NL

# 9
H2 "9. Achado Adicional — Divergencia entre Taxa Declarada e TIR Real"
Ital "O banco declara 9,99% a.m., mas o PMT cobrado (R$ 636,64) implica TIR de 10,9352% a.m. — 0,94 p.p. acima do declarado. A diferenca de R$ 40,05/parcela x 21 = R$ 841,04 cobrados a mais do que o PRICE a 9,99% produziria. Causa provavel: capitalizacao dos juros da carencia de 48 dias sem informacao clara ao consumidor. Fundamento: CDC art. 6, III; Res. BACEN 3.517/2007; CDC art. 51, IV."
NL

Para "Calculos elaborados com base no extrato Midway emitido em 17/06/2026."
Para "Serie BACEN 25479 consultada em 17/06/2026 via API oficial do Banco Central."
Para "Sujeito a conferencia por perito judicial."

# Salvar
$doc.SaveAs([ref]$docPath, [ref]16)
$doc.Close()
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

if (Test-Path $docPath) { "CONCLUIDO: $docPath" } else { "ERRO: arquivo nao criado" }
