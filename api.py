# ce fichier permet d'utiliser l'api fastapi
import io
import json
import os
import language_tool_python
from docx import Document
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from corrections import corriger_document
from ia import corriger_avec_ia

TAILLE_MAX = 5 * 1024 * 1024

with open("fautes_orthographes/errors.json", "r", encoding="utf-8") as f:
    data = json.load(f)

with open("fautes_orthographes/ignorer.json", "r", encoding="utf-8") as f:
    ignores = set(json.load(f))

tl = None


def get_outil():
    global tl
    if tl is None:
        if os.getenv("LT_MODE") == "public":
            tl = language_tool_python.LanguageToolPublicAPI("fr")
        else:
            tl = language_tool_python.LanguageTool("fr")
    return tl


app = FastAPI()
origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,              # QUI a le droit d'appeler l'API
    allow_methods=["*"],                # avec quelles méthodes (GET, POST…)
    allow_headers=["*"],                # avec quels en-têtes
    expose_headers=["X-Correction"],
)


@app.get("/")
def read_root():
    return {"message": "Le correcteur fonctionne"}


@app.post("/corriger")
async def creer_fichier(file: UploadFile):
    octets = await file.read()

    if len(octets) > TAILLE_MAX:
        raise HTTPException(status_code=413, detail="Fichier trop volumineux (5 Mo max)")

    try:
        document = Document(io.BytesIO(octets))  # ouvre contenu doc
    except Exception:
        raise HTTPException(status_code=400, detail="Ce fichier n'est pas un document Word valide")

    # 1er choix : l'IA (corrige même les fautes écrites au son)
    statut = "ia"
    try:
        corriger_avec_ia(document, ignores)
    except Exception as e:
        print("IA INDISPONIBLE :", e)
        document = Document(io.BytesIO(octets))      # on repart du fichier d'origine
        # plan B : dictionnaire + LanguageTool
        statut = "complete"
        try:
            corriger_document(document, data, ignores, get_outil())
        except Exception as e:
            print("LANGUAGETOOL INDISPONIBLE :", e)
            document = Document(io.BytesIO(octets))
            corriger_document(document, data, ignores, None)
            statut = "partielle"

    sortie = io.BytesIO()
    document.save(sortie)
    sortie.seek(0)
    nom = file.filename.replace(".docx", "_corrige.docx")
    return StreamingResponse(  # 7. renvoyer le fichier
        sortie,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{nom}"',
            "X-Correction": statut,
        },
    )