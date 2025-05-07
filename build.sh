#!/bin/bash

# Atualizar pip
python -m pip install --upgrade pip

# Instalar dependências
pip install -r requirements.txt

# Instalar versão específica do Streamlit
pip install streamlit==1.31.1 