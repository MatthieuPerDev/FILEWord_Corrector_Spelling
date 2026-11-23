# image de départ sur laquelle mon conteneur va être construit
FROM python:3.13-slim

# Java pour LanguageTool (mode local)
RUN apt-get update && apt-get install -y default-jre-headless && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Mémoire de Java limitée pour tenir dans les 512 Mo de Render gratuit (Java ~320 Mo + API ~60 Mo)
ENV JAVA_TOOL_OPTIONS="-Xms32m -Xmx224m -XX:+UseSerialGC -XX:TieredStopAtLevel=1 -Xss512k -XX:MaxMetaspaceSize=96m -XX:ReservedCodeCacheSize=32m"

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# téléchargement de LanguageTool pendant le build
RUN python -c "import language_tool_python; language_tool_python.LanguageTool('fr').close()"

COPY . .
EXPOSE 8000

# Render impose son port via PORT ; en local on garde 8000
CMD ["sh", "-c", "python -m uvicorn api:app --host 0.0.0.0 --port ${PORT:-8000}"]