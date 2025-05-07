#!/bin/bash

# Criar diretório para configuração do Streamlit
mkdir -p ~/.streamlit

# Configurar o arquivo config.toml
cat > ~/.streamlit/config.toml << EOL
[server]
maxUploadSize = 200
enableCORS = false
enableXsrfProtection = false
port = $PORT

[browser]
gatherUsageStats = false
EOL

# Configurar variáveis de ambiente
export PYTHONUNBUFFERED=1
export PORT=8501

echo "\
[general]\n\
email = \"\"\n\
" > ~/.streamlit/credentials.toml 