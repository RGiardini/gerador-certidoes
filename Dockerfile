FROM python:3.10-slim

# Instala o LibreOffice de forma enxuta para converter os PDFs
RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    default-jre \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copia o arquivo de requisitos e instala as bibliotecas do Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia todo o seu código para o servidor
COPY . .

# Expõe a porta que o Streamlit usa
EXPOSE 8501

# Comando para rodar o Streamlit (MUDE O NOME DO ARQUIVO ABAIXO SE NECESSÁRIO)
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
