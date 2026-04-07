@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "& {

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Assistente Profissional de Orçamento'
$form.Size = New-Object System.Drawing.Size(600,500)
$form.StartPosition = 'CenterScreen'
$form.BackColor = '#f4f4f4'

# ===== Título =====
$title = New-Object System.Windows.Forms.Label
$title.Text = 'Assistente de Orçamento'
$title.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(150,20)
$form.Controls.Add($title)

# ===== BOTÃO MANUAL =====
$btnManual = New-Object System.Windows.Forms.Button
$btnManual.Text = 'Digitar Manualmente'
$btnManual.Size = New-Object System.Drawing.Size(200,50)
$btnManual.Location = New-Object System.Drawing.Point(200,80)
$form.Controls.Add($btnManual)

# ===== BOTÃO ANEXO =====
$btnArquivo = New-Object System.Windows.Forms.Button
$btnArquivo.Text = 'Anexar Orçamento'
$btnArquivo.Size = New-Object System.Drawing.Size(200,50)
$btnArquivo.Location = New-Object System.Drawing.Point(200,150)
$form.Controls.Add($btnArquivo)

# ===== FORMULÁRIO =====
$panel = New-Object System.Windows.Forms.Panel
$panel.Size = New-Object System.Drawing.Size(500,250)
$panel.Location = New-Object System.Drawing.Point(50,220)
$panel.Visible = $false
$form.Controls.Add($panel)

function CriarCampo($texto,$y){
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $texto
    $label.Location = New-Object System.Drawing.Point(10,$y)
    $panel.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Size = New-Object System.Drawing.Size(300,25)
    $textbox.Location = New-Object System.Drawing.Point(150,$y)
    $panel.Controls.Add($textbox)

    return $textbox
}

$produto = CriarCampo 'Produto:' 10
$ncm = CriarCampo 'NCM:' 40
$ipi = CriarCampo 'IPI (%):' 70

# ===== FRETE =====
$labelFrete = New-Object System.Windows.Forms.Label
$labelFrete.Text = 'Frete:'
$labelFrete.Location = New-Object System.Drawing.Point(10,100)
$panel.Controls.Add($labelFrete)

$comboFrete = New-Object System.Windows.Forms.ComboBox
$comboFrete.Items.AddRange(@('FOB','CIF'))
$comboFrete.Location = New-Object System.Drawing.Point(150,100)
$panel.Controls.Add($comboFrete)

# ===== FILIAL =====
$labelFilial = New-Object System.Windows.Forms.Label
$labelFilial.Text = 'Filial:'
$labelFilial.Location = New-Object System.Drawing.Point(10,130)
$panel.Controls.Add($labelFilial)

$comboFilial = New-Object System.Windows.Forms.ComboBox
$comboFilial.Items.Add('Natal - 18.217.682/0004-05')
$comboFilial.Items.Add('Pernambuco - 18.217.682/0001-54')
$comboFilial.Location = New-Object System.Drawing.Point(150,130)
$comboFilial.Size = New-Object System.Drawing.Size(300,25)
$panel.Controls.Add($comboFilial)

# ===== BOTÃO SALVAR =====
$btnSalvar = New-Object System.Windows.Forms.Button
$btnSalvar.Text = 'Confirmar'
$btnSalvar.Location = New-Object System.Drawing.Point(150,170)
$panel.Controls.Add($btnSalvar)

# ===== EVENTOS =====

$btnManual.Add_Click({
    $panel.Visible = $true
})

$btnArquivo.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = 'Arquivos|*.xlsx;*.csv;*.pdf'

    if($fileDialog.ShowDialog() -eq 'OK'){
        [System.Windows.Forms.MessageBox]::Show('Arquivo selecionado: ' + $fileDialog.FileName)
    }
})

$btnSalvar.Add_Click({
    $resumo = "Produto: $($produto.Text)`nNCM: $($ncm.Text)`nIPI: $($ipi.Text)`nFrete: $($comboFrete.Text)`nFilial: $($comboFilial.Text)"
    [System.Windows.Forms.MessageBox]::Show($resumo,'Dados preenchidos')
})

$form.ShowDialog()

}"
