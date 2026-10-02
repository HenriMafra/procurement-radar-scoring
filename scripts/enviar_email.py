#!/usr/bin/env python3
import json
import sys
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

if len(sys.argv) < 2:
    print("Uso: python3 enviar_email.py <json_path>")
    sys.exit(1)

json_path = sys.argv[1]

try:
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    dests = data.get("to", "henri.mafra@grupoenterprisecore.com")
    subject = data.get("subject", "[B2G-RADAR-ENTERPRISECORE] Alerta de Movimentação")
    html_body = data.get("html", "")
    
    # Carrega config de SMTP se existir em config/smtp.json ou ambiente
    smtp_file = os.path.join(os.path.dirname(__file__), "..", "config", "smtp.json")
    smtp_cfg = {}
    if os.path.exists(smtp_file):
        try:
            with open(smtp_file, "r", encoding="utf-8") as sf:
                smtp_cfg = json.load(sf)
        except Exception:
            pass

    host = smtp_cfg.get("host", "smtp.office365.com")
    port = int(smtp_cfg.get("port", 587))
    user = smtp_cfg.get("user", "")
    password = smtp_cfg.get("password", "")
    sender = smtp_cfg.get("sender", user or "radar-alertas@grupoenterprisecore.com")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = dests
    msg["X-B2G-System"] = "ENTERPRISECORE-RADAR-247"
    msg["X-B2G-Filter-Key"] = "B2G-RADAR-ENTERPRISECORE"

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if user and password:
        with smtplib.SMTP(host, port, timeout=15) as server:
            server.starttls()
            server.login(user, password)
            to_list = [e.strip() for e in dests.replace(";", ",").split(",") if e.strip()]
            server.sendmail(sender, to_list, msg.as_string())
        print(f"✅ Alerta enviado via SMTP ({host}) para: {dests}")
    else:
        # Grava em log de despachos e stdout
        log_dir = os.path.join(os.path.dirname(__file__), "..", "data", "alertas_enviados.log")
        with open(log_dir, "a", encoding="utf-8") as lf:
            lf.write(f"[{subject}] -> Para: {dests}\n")
        print(f"ℹ️ Alerta registrado para envio: {subject} -> {dests}")

except Exception as e:
    print(f"❌ Erro ao enviar e-mail: {e}", file=sys.stderr)
    sys.exit(1)
