@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
mode con: cols=110 lines=35
color 1F
set "APPNAME=Assistente Profissional de Orcamento"
set "TEMPPS=%temp%\orcamento_gui_%random%%random%.ps1"
set "RESULT=%temp%\orcamento_result_%random%%random%.txt"

goto :MENU

:MENU
cls
echo ==============================================================================================================
echo                                        %APPNAME%
echo ==============================================================================================================
echo.
echo   [1] Digitar informacoes manualmente
echo   [2] Anexar orcamento para preenchimento automatico
echo   [0] Sair
echo.
set /p "OPCAO=Escolha uma opcao e pressione ENTER: "

if "%OPCAO%"=="1" goto :MANUAL
if "%OPCAO%"=="2" goto :AUTO
if "%OPCAO%"=="0" goto :END

echo.
echo Opcao invalida.
timeout /t 2 >nul
goto :MENU

:MANUAL
call :WRITE_PS_MANUAL
powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMPPS%" "%RESULT%"
if not exist "%RESULT%" (
    echo.
    echo Operacao cancelada pelo usuario.
    timeout /t 2 >nul
    goto :CLEAN
)
call :SHOW_RESULT
set /p "NEXT=Deseja voltar ao menu principal? [S/N]: "
if /I "%NEXT%"=="S" goto :CLEAN
if /I "%NEXT%"=="N" goto :END
goto :CLEAN

:AUTO
call :WRITE_PS_AUTO
powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMPPS%" "%RESULT%"
if not exist "%RESULT%" (
    echo.
    echo Nenhum arquivo foi processado.
    timeout /t 2 >nul
    goto :CLEAN
)
call :SHOW_RESULT
set /p "NEXT=Deseja voltar ao menu principal? [S/N]: "
if /I "%NEXT%"=="S" goto :CLEAN
if /I "%NEXT%"=="N" goto :END
goto :CLEAN

:SHOW_RESULT
cls
echo ==============================================================================================================
echo                                         RESULTADO FINAL
echo ==============================================================================================================
echo.
type "%RESULT%"
echo.
echo --------------------------------------------------------------------------------------------------------------
echo Observacoes sobre anexos:
echo - CSV: preenchimento automatico imediato.
echo - XLSX: preenchimento automatico se o Microsoft Excel estiver instalado.
echo - PDF, imagem ou layout fora do padrao: o BAT identifica o arquivo, mas a extracao robusta exige apoio adicional.
echo - Padrao recomendado de colunas para automacao: Produto, NCM, IPI, Frete, Filial.
echo --------------------------------------------------------------------------------------------------------------
echo.
exit /b

:WRITE_PS_MANUAL
> "%TEMPPS%" (
    echo Add-Type -AssemblyName System.Windows.Forms
    echo Add-Type -AssemblyName System.Drawing
    echo $out = $args[0]
    echo $form = New-Object System.Windows.Forms.Form
    echo $form.Text = 'Cadastro Manual de Orcamento'
    echo $form.Size = New-Object System.Drawing.Size^(700,420^)
    echo $form.StartPosition = 'CenterScreen'
    echo $form.FormBorderStyle = 'FixedDialog'
    echo $form.MaximizeBox = $false
    echo $font = New-Object System.Drawing.Font^('Segoe UI',10^)
    echo $form.Font = $font
    echo function Add-Label^($text,$x,$y^){ $l = New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point^($x,$y^); $l.Size=New-Object System.Drawing.Size^(180,25^); $form.Controls.Add^($l^) }
    echo function Add-TextBox^($x,$y,$w^){ $t = New-Object System.Windows.Forms.TextBox; $t.Location=New-Object System.Drawing.Point^($x,$y^); $t.Size=New-Object System.Drawing.Size^($w,25^); $form.Controls.Add^($t^); return $t }
    echo Add-Label 'Nome do Produto:' 20 30
    echo $txtProduto = Add-TextBox 210 28 440
    echo Add-Label 'NCM:' 20 70
    echo $txtNCM = Add-TextBox 210 68 180
    echo Add-Label 'IPI ^(%%^):' 20 110
    echo $txtIPI = Add-TextBox 210 108 120
    echo Add-Label 'Frete:' 20 150
    echo $cmbFrete = New-Object System.Windows.Forms.ComboBox
    echo $cmbFrete.Location = New-Object System.Drawing.Point^(210,148^)
    echo $cmbFrete.Size = New-Object System.Drawing.Size^(120,25^)
    echo $cmbFrete.DropDownStyle = 'DropDownList'
    echo [void]$cmbFrete.Items.AddRange^(@^('FOB','CIF'^)^)
    echo $cmbFrete.SelectedIndex = 0
    echo $form.Controls.Add^($cmbFrete^)
    echo Add-Label 'Filial de compra:' 20 190
    echo $cmbFilial = New-Object System.Windows.Forms.ComboBox
    echo $cmbFilial.Location = New-Object System.Drawing.Point^(210,188^)
    echo $cmbFilial.Size = New-Object System.Drawing.Size^(440,25^)
    echo $cmbFilial.DropDownStyle = 'DropDownList'
    echo [void]$cmbFilial.Items.AddRange^(@^('Natal - CNPJ: 18.217.682/0004-05','Pernambuco - CNPJ: 18.217.682/0001-54'^)^)
    echo $cmbFilial.SelectedIndex = 0
    echo $form.Controls.Add^($cmbFilial^)
    echo $info = New-Object System.Windows.Forms.Label
    echo $info.Location = New-Object System.Drawing.Point^(20,240^)
    echo $info.Size = New-Object System.Drawing.Size^(630,60^)
    echo $info.Text = 'Preencha os campos e clique em Confirmar. O arquivo BAT gera um resumo pronto para uso ou integracao com outra automacao.'
    echo $form.Controls.Add^($info^)
    echo $btnOk = New-Object System.Windows.Forms.Button
    echo $btnOk.Text = 'Confirmar'
    echo $btnOk.Location = New-Object System.Drawing.Point^(430,315^)
    echo $btnOk.Size = New-Object System.Drawing.Size^(105,32^)
    echo $btnCancel = New-Object System.Windows.Forms.Button
    echo $btnCancel.Text = 'Cancelar'
    echo $btnCancel.Location = New-Object System.Drawing.Point^(545,315^)
    echo $btnCancel.Size = New-Object System.Drawing.Size^(105,32^)
    echo $form.Controls.Add^($btnOk^)
    echo $form.Controls.Add^($btnCancel^)
    echo $btnCancel.Add_Click^({ $form.Close^(^) })
    echo $btnOk.Add_Click^({
    echo   if ^([string]::IsNullOrWhiteSpace^($txtProduto.Text^) -or [string]::IsNullOrWhiteSpace^($txtNCM.Text^) -or [string]::IsNullOrWhiteSpace^($txtIPI.Text^)^) {
    echo     [System.Windows.Forms.MessageBox]::Show^('Preencha Produto, NCM e IPI.','Validacao'^) ^| Out-Null
    echo     return
    echo   }
    echo   $filial = $cmbFilial.SelectedItem
    echo   $conteudo = @(
    echo     'MODO: DIGITACAO MANUAL',
    echo     'Produto: ' + $txtProduto.Text,
    echo     'NCM: ' + $txtNCM.Text,
    echo     'IPI: ' + $txtIPI.Text + '%%',
    echo     'Frete: ' + $cmbFrete.SelectedItem,
    echo     'Filial: ' + $filial,
    echo     '',
    echo     'Status: dados capturados com sucesso.'
    echo   )
    echo   Set-Content -Path $out -Value $conteudo -Encoding UTF8
    echo   $form.Close^(^)
    echo })
    echo [void]$form.ShowDialog^(^)
)
exit /b

:WRITE_PS_AUTO
> "%TEMPPS%" (
    echo Add-Type -AssemblyName System.Windows.Forms
    echo Add-Type -AssemblyName Microsoft.VisualBasic
    echo $out = $args[0]
    echo $ofd = New-Object System.Windows.Forms.OpenFileDialog
    echo $ofd.Title = 'Selecione o orcamento'
    echo $ofd.Filter = 'Arquivos suportados ^(*.csv;*.xlsx;*.xls^)|*.csv;*.xlsx;*.xls^|CSV ^(*.csv^)|*.csv^|Excel ^(*.xlsx;*.xls^)|*.xlsx;*.xls^|Todos os arquivos ^(*.*^)|*.*'
    echo $ofd.Multiselect = $false
    echo if ^($ofd.ShowDialog^(^) -ne 'OK'^) { exit }
    echo $arquivo = $ofd.FileName
    echo $ext = [System.IO.Path]::GetExtension^($arquivo^).ToLower^(^)
    echo $linhas = New-Object System.Collections.Generic.List[string]
    echo $linhas.Add^('MODO: PREENCHIMENTO AUTOMATICO'^)
    echo $linhas.Add^('Arquivo anexado: ' + $arquivo^)
    echo $linhas.Add^(''^)
    echo if ^($ext -eq '.csv'^) {
    echo   try {
    echo     $dados = Import-Csv -Path $arquivo
    echo     if ^(-not $dados -or $dados.Count -eq 0^) { throw 'Arquivo CSV vazio.' }
    echo     $primeiro = $dados[0]
    echo     $produto = $primeiro.Produto
    echo     if ^(-not $produto^) { $produto = $primeiro.'Nome do Produto' }
    echo     $ncm = $primeiro.NCM
    echo     $ipi = $primeiro.IPI
    echo     $frete = $primeiro.Frete
    echo     $filial = $primeiro.Filial
    echo     if ^(-not $produto^) { $produto = '[nao identificado automaticamente]' }
    echo     if ^(-not $ncm^) { $ncm = '[nao identificado automaticamente]' }
    echo     if ^(-not $ipi^) { $ipi = '[nao identificado automaticamente]' }
    echo     if ^(-not $frete^) { $frete = '[nao identificado automaticamente]' }
    echo     if ^(-not $filial^) { $filial = '[nao identificado automaticamente]' }
    echo     $linhas.Add^('Produto: ' + $produto^)
    echo     $linhas.Add^('NCM: ' + $ncm^)
    echo     $linhas.Add^('IPI: ' + $ipi^)
    echo     $linhas.Add^('Frete: ' + $frete^)
    echo     $linhas.Add^('Filial: ' + $filial^)
    echo     $linhas.Add^(''^)
    echo     $linhas.Add^('Status: leitura automatica concluida via CSV.'^)
    echo     $linhas.Add^('Padrao recomendado de colunas: Produto, NCM, IPI, Frete, Filial.'^)
    echo   } catch {
    echo     $linhas.Add^('Falha ao ler CSV: ' + $_.Exception.Message^)
    echo   }
    echo } elseif ^($ext -eq '.xlsx' -or $ext -eq '.xls'^) {
    echo   try {
    echo     $excel = New-Object -ComObject Excel.Application
    echo     $excel.Visible = $false
    echo     $wb = $excel.Workbooks.Open^($arquivo^)
    echo     $ws = $wb.Worksheets.Item^(1^)
    echo     $used = $ws.UsedRange.Value2
    echo     $headers = @{}
    echo     for ^($c=1; $c -le $used.GetLength^(1^); $c++^) {
    echo       $nome = [string]$used[1,$c]
    echo       if ^($nome^) { $headers[$nome.Trim^(^).ToLower^(^)] = $c }
    echo     }
    echo     function GetVal^($map,$data,$row,$names^) {
    echo       foreach ^($n in $names^) { if ^($map.ContainsKey^($n^)^) { return [string]$data[$row,$map[$n]] } }
    echo       return $null
    echo     }
    echo     $produto = GetVal $headers $used 2 @^('produto','nome do produto'^)
    echo     $ncm = GetVal $headers $used 2 @^('ncm'^)
    echo     $ipi = GetVal $headers $used 2 @^('ipi'^)
    echo     $frete = GetVal $headers $used 2 @^('frete'^)
    echo     $filial = GetVal $headers $used 2 @^('filial'^)
    echo     if ^(-not $produto^) { $produto = '[nao identificado automaticamente]' }
    echo     if ^(-not $ncm^) { $ncm = '[nao identificado automaticamente]' }
    echo     if ^(-not $ipi^) { $ipi = '[nao identificado automaticamente]' }
    echo     if ^(-not $frete^) { $frete = '[nao identificado automaticamente]' }
    echo     if ^(-not $filial^) { $filial = '[nao identificado automaticamente]' }
    echo     $linhas.Add^('Produto: ' + $produto^)
    echo     $linhas.Add^('NCM: ' + $ncm^)
    echo     $linhas.Add^('IPI: ' + $ipi^)
    echo     $linhas.Add^('Frete: ' + $frete^)
    echo     $linhas.Add^('Filial: ' + $filial^)
    echo     $linhas.Add^(''^)
    echo     $linhas.Add^('Status: leitura automatica concluida via Excel.'^)
    echo     $linhas.Add^('Padrao recomendado de colunas: Produto, NCM, IPI, Frete, Filial.'^)
    echo     $wb.Close^($false^)
    echo     $excel.Quit^(^)
    echo     [System.Runtime.Interopservices.Marshal]::ReleaseComObject^($ws^) ^| Out-Null
    echo     [System.Runtime.Interopservices.Marshal]::ReleaseComObject^($wb^) ^| Out-Null
    echo     [System.Runtime.Interopservices.Marshal]::ReleaseComObject^($excel^) ^| Out-Null
    echo   } catch {
    echo     $linhas.Add^('Falha ao ler Excel automaticamente: ' + $_.Exception.Message^)
    echo     $linhas.Add^('Dica: use CSV ou mantenha colunas padrao na primeira linha.'^)
    echo   }
    echo } else {
    echo   $linhas.Add^('Formato anexado nao possui leitura automatica completa neste BAT.'^)
    echo   $linhas.Add^('Use preferencialmente CSV ou XLSX com cabecalho: Produto, NCM, IPI, Frete, Filial.'^)
    echo }
    echo Set-Content -Path $out -Value $linhas -Encoding UTF8
)
exit /b

:CLEAN
if exist "%TEMPPS%" del /f /q "%TEMPPS%" >nul 2>nul
if exist "%RESULT%" del /f /q "%RESULT%" >nul 2>nul
goto :MENU

:END
if exist "%TEMPPS%" del /f /q "%TEMPPS%" >nul 2>nul
if exist "%RESULT%" del /f /q "%RESULT%" >nul 2>nul
endlocal
exit /b
