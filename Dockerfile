FROM python:3.10-slim

WORKDIR /app

# Copia o arquivo de requisitos e instala as bibliotecas do Python (incluindo o reportlab)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia todo o seu código para o servidor
COPY . .

# Expõe a porta que o Streamlit usa
EXPOSE 8501

# Comando para rodar o Streamlit
CMD sh -c "streamlit run app.py --server.port=${PORT:-8501} --server.address=0.0.0.0 --server.fileWatcherType=none"