#!/bin/bash
# Navega até o diretório onde o script está localizado
cd "$(dirname "$0")"

echo "=================================================="
echo "    Iniciando NutriFill AI - Servidor Local       "
echo "=================================================="

# Verifica se o Python 3 está instalado
if ! command -v python3 &> /dev/null
then
    echo "[ERRO] Python 3 não foi encontrado no seu Mac!"
    echo "Por favor, faça o download e instale o Python 3 em:"
    echo "https://www.python.org/downloads/macos/"
    echo ""
    echo "Pressione qualquer tecla para sair..."
    read -n 1
    exit 1
fi

# Verifica e instala as dependências necessárias
echo "Verificando e instalando dependências necessárias..."
python3 -m pip install --upgrade pip
python3 -m pip install flask pypdf openpyxl google-genai pydantic

# Abre o navegador padrão na página local
echo "Abrindo o navegador em http://127.0.0.1:5000 ..."
sleep 1
open "http://127.0.0.1:5000"

# Inicia o servidor Flask local
python3 app.py
