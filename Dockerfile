FROM python:3.10-slim

WORKDIR /app

# Copia o arquivo de requisitos e instala as bibliotecas do Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia todo o seu código para o servidor
COPY . .

# Expõe a porta padrão (o Render sobrescreve via variável $PORT)
EXPOSE 8501

# Comando correto usando shell form para o Render expandir a variável $PORT dinamicamente
CMD streamlit run app.py --server.port=$PORT --server.address=0.0.0.0 --server.fileWatcherType=none