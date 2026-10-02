param(
  [string]$JsonPath,
  [string]$OutPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if(-not (Test-Path -LiteralPath $JsonPath)){ throw "Arquivo JSON nao encontrado: $JsonPath" }

$raw = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$to    = [string]$raw.to
$cc    = [string]$raw.cc
$items = @($raw.items)

if($items.Count -eq 0){ throw "Nenhum item fornecido para gerar .OFT" }

$outlook = $null
$tmpDir = Join-Path $env:TEMP "enterprisecore_oft_batch_$(Get-Random)"
try {
    $outlook = New-Object -ComObject Outlook.Application
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

    # Cada item pode indicar em qual pasta do .zip deve cair (ex.: "Radar de
    # Publicacoes" para o e-mail unico da regiao, "Individuais" para o resto).
    # Sem "pasta" informada, cai direto na raiz do zip.
    $idx = 1
    foreach($x in $items){
        $pasta = [string]$x.pasta
        $destDir = $tmpDir
        if(-not [string]::IsNullOrWhiteSpace($pasta)){
            $destDir = Join-Path $tmpDir $pasta
            if(-not (Test-Path -LiteralPath $destDir)){ New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        }

        $cleanOrgao = ($x.orgao -replace '[^a-zA-Z0-9]', '_').Trim('_')
        if([string]::IsNullOrWhiteSpace($cleanOrgao)){ $cleanOrgao = "Oportunidade_$idx" }
        if($cleanOrgao.Length -gt 35){ $cleanOrgao = $cleanOrgao.Substring(0, 35) }

        $numConlic = [string]$x.numeroConlicitacao
        $fileName = "Alerta-ENTERPRISECORE-$cleanOrgao-$numConlic.oft"
        $itemOftPath = Join-Path $destDir $fileName

        $mail = $null
        try {
            $mail = $outlook.CreateItem(0)
            if($to){ $mail.To = $to }
            if($cc){ $mail.CC = $cc }
            if($x.subject){ $mail.Subject = $x.subject }
            if($x.html){ $mail.HTMLBody = $x.html }

            $mail.SaveAs($itemOftPath, 2) # 2 = olTemplate
            $mail.Close(1)
        } finally {
            if($mail){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null }
        }
        $idx++
    }

    if(Test-Path -LiteralPath $OutPath){ Remove-Item -LiteralPath $OutPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpDir, $OutPath)
    Write-Host "OFT_ZIP_OK"
} catch {
    # exit 1 para o servidor conseguir mostrar o motivo real da falha
    Write-Error "ERRO_OFT_BATCH: $($_.Exception.Message)"
    exit 1
} finally {
    if($outlook){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    if(Test-Path -LiteralPath $tmpDir){ Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
