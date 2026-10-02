#!/bin/bash
set -e

echo "=== Instalando ENTERPRISECORE Radar na Nuvem (Ubuntu 24.04 OCI) ==="

# 1. Dependências do Sistema
apt-get update -y
apt-get install -y nodejs npm python3 python3-pip curl

# 2. Diretório de Instalação
INSTALL_DIR="/opt/enterprisecore-radar"
mkdir -p $INSTALL_DIR
tar -xzf /tmp/enterprisecore-radar-cloud-pkg.tar.gz -C $INSTALL_DIR

cd $INSTALL_DIR

# 3. Permissões
chown -R ubuntu:ubuntu $INSTALL_DIR

# 4. Cria Serviço Systemd
cat << 'EOF' > /etc/systemd/system/enterprisecore-radar.service
[Unit]
Description=ENTERPRISECORE Radar - Monitoramento de Licitacoes 24/7
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/opt/enterprisecore-radar
ExecStart=/usr/bin/node /opt/enterprisecore-radar/app/server.js
Restart=always
RestartSec=10
Environment=NODE_ENV=production
Environment=PORT=8790

[Install]
WantedBy=multi-user.target
EOF

# 5. Ativa e Inicia o Serviço
systemctl daemon-reload
systemctl enable enterprisecore-radar
systemctl restart enterprisecore-radar

echo "=== ENTERPRISECORE Radar Ativo e Rodando 24/7 via Systemd! ==="
systemctl status enterprisecore-radar --no-pager
