# ce fichier permet d'utiliser l'api fastapi
import io
from docx import Document
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from corrections import corriger_document, charger_json, est_public, outil_public, get_outil as demarrer_languagetool
from ia import corriger_avec_ia

TAILLE_MAX = 5 * 1024 * 1024

data = charger_json("errors.json", {})
ignores = set(charger_json("ignorer.json", []))

tl = None
tl_essaye = False


def get_outil():
    """LanguageTool démarré une seule fois ; None s'il est indisponible (-> dictionnaires seuls)."""
    global tl, tl_essaye
    if not tl_essaye:
        tl_essaye = True
        tl = demarrer_languagetool()
    return tl


app = FastAPI()
origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,              # QUI a le droit d'appeler l'API
    allow_methods=["*"],                # avec quelles méthodes (GET, POST…)
    allow_headers=["*"],                # avec quels en-têtes
    expose_headers=["X-Correction", "X-Nb-Corrections"],
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
    global tl
    statut = "ia"
    try:
        nb = corriger_avec_ia(document, ignores)
    except Exception as e:
        print("IA INDISPONIBLE :", e)
        document = Document(io.BytesIO(octets))      # on repart du fichier d'origine
        # plan B : dictionnaire + LanguageTool
        infos = {}
        _, nb = corriger_document(document, data, ignores, get_outil(), infos)
        statut = "complete" if infos["mode"] in ("local", "public") else "partielle"
        if infos["mode"] == "public" and tl is not None and not est_public(tl):
            tl = outil_public()                      # le LanguageTool local a planté -> on garde l'en ligne
    print(f"CORRECTION TERMINÉE : {nb} correction(s), mode {statut}")

    sortie = io.BytesIO()
    document.save(sortie)
    sortie.seek(0)
    nom = (file.filename or "document.docx").replace(".docx", "_corrige.docx")
    return StreamingResponse(  # 7. renvoyer le fichier
        sortie,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{nom}"',
            "X-Correction": statut,
            "X-Nb-Corrections": str(nb),
        },
    )