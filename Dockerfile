#image de depart sur lequel mon conteneur va etre construit
FROM python:3.13-slim

RUN apt-get update && apt-get install -y default-jre-headless && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN python -c "import language_tool_python; language_tool_python.LanguageTool('fr').close()"
COPY . .
EXPOSE 8000
CMD ["python", "-m", "uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]