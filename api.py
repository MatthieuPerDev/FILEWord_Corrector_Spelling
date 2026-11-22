# ce fichier permet d'utiliser l'api fastapi
import io
import json
import os                                                    # ← lire les variables d'environnement
import language_tool_python
from docx import Document
from fastapi import FastAPI, UploadFile, HTTPException       # ← HTTPException pour les erreurs
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from corrections import corriger_document

TAILLE_MAX = 5 * 1024 * 1024                                 # ← 5 Mo (1 Mo = 1024 Ko = 1024*1024 octets)

with open("fautes_orthographes/errors.json", "r", encoding="utf-8") as f:
    data = json.load(f)

with open("fautes_orthographes/ignorer.json", "r", encoding="utf-8") as f:
    ignores = set(json.load(f))
# LT_MODE=public sur Render (peu de RAM) ; en local : LanguageTool complet avec Java
if os.getenv("LT_MODE") == "public":
    tl = language_tool_python.LanguageToolPublicAPI("fr")
else:
    tl = language_tool_python.LanguageTool("fr")
app = FastAPI()

# adresses autorisées : localhost par défaut, l'adresse Render en ligne
origins = os.getenv("FRONTEND_URL", "http://localhost:5173").split(",")   # ←

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,              # QUI a le droit d'appeler l'API
    allow_methods=["*"],                # avec quelles méthodes (GET, POST…)
    allow_headers=["*"],                # avec quels en-têtes
)

@app.get("/")
def read_root():
    return {"message": "Le correcteur fonctionne"}

@app.post("/corriger")
async def creer_fichier(file: UploadFile):
    octets = await file.read()

    if len(octets) > TAILLE_MAX:                              # ← sécurité : refuser les gros fichiers
        raise HTTPException(status_code=413, detail="Fichier trop volumineux (5 Mo max)")

    try:                                                      # ← fichier pas Word → erreur propre (400)
        document = Document(io.BytesIO(octets))  # ouvre contenu doc
    except Exception:
        raise HTTPException(status_code=400, detail="Ce fichier n'est pas un document Word valide")

    try:
        corriger_document(document, data, ignores, tl)
    except Exception as e:
        print("ERREUR CORRECTION :", e)  # visible dans le terminal de l'API
        raise HTTPException(status_code=502, detail=f"Le service de correction ne répond pas : {e}")
    sortie = io.BytesIO()
    document.save(sortie)
    sortie.seek(0)
    nom = file.filename.replace(".docx", "_corrige.docx")
    return StreamingResponse(  # renvoyer le fichier
        sortie,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{nom}"'},
    )