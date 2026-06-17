# Usa uma imagem oficial e leve do Python
FROM python:3.11-slim

# Define o diretório de trabalho dentro do container
WORKDIR /app

# Copia o arquivo de dependências primeiro para aproveitar o cache do Docker
COPY requirements.txt .

# Instala as dependências sem armazenar cache desnecessário na imagem
RUN pip install --no-cache-dir -r requirements.txt

# Copia o restante do código da aplicação (seu main.py)
COPY . .

# Expõe a porta padrão utilizada pelo Streamlit
EXPOSE 8501

# Comando de inicialização forçando o endereço de rede correto para serviços em nuvem
CMD ["streamlit", "run", "main.py", "--server.port=8501", "--server.address=0.0.0.0"]