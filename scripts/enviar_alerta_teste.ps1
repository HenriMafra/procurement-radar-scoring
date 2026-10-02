[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# REGRA ESTRITA: E-mails de teste vao APENAS para henri.mafra@grupoenterprisecore.com com prefixo "TESTE - "
$To = "henri.mafra@grupoenterprisecore.com"
$Subject = "TESTE - [B2G-RADAR-ENTERPRISECORE] Validacao de Alerta e Layout"

$HtmlBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body style="margin:0;padding:0;background:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f1f5f9;padding:24px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="680" style="max-width:680px;background:#ffffff;border:1px solid #cbd5e1;border-radius:8px;overflow:hidden;box-shadow:0 4px 14px rgba(0,0,0,0.06);">
          <tr>
            <td style="padding:18px 24px;background:#ffffff;border-bottom:2px solid #e2e8f0;">
              <div style="font-size:13px;font-weight:800;color:#0f172a;letter-spacing:-0.01em;text-transform:uppercase;">ENTERPRISECORE RADAR &bull; DISPARO DE TESTE</div>
              <div style="font-size:11px;color:#64748b;margin-top:2px;">Assunto do Teste: <span style="font-family:monospace;font-weight:700;">TESTE - [B2G-RADAR-ENTERPRISECORE]</span></div>
            </td>
          </tr>
          <tr>
            <td style="padding:24px 28px;background:#ffffff;">
              <div style="background:#fff1f2;border-left:5px solid #e11d48;padding:16px 18px;border-radius:4px;margin-bottom:22px;">
                <div style="color:#be123c;font-size:13px;font-weight:900;text-transform:uppercase;">
                  TESTE DE LAYOUT E DISPARO DE E-MAIL
                </div>
                <div style="margin-top:8px;font-size:13.5px;color:#1e293b;">
                  Este e-mail e um teste de sistema. Disparado exclusivamente para henri.mafra@grupoenterprisecore.com.
                </div>
              </div>
            </td>
          </tr>
          <tr>
            <td style="padding:14px 24px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11.5px;color:#64748b;text-align:center;">
              ENTERPRISECORE Radar &bull; Ambiente de Testes &bull; Enterprise IT Group
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"@

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    $mail.HTMLBody = $HtmlBody
    $mail.Send()
    Write-Output "OK: E-mail de teste enviado EXCLUSIVAMENTE para $To"
} catch {
    Write-Error "Erro: $_"
}
