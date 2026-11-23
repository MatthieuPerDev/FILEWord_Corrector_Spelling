# ce fichier permet d'utiliser l'api fastapi
import io
import threading
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
verrou = threading.Lock()


def get_outil(attente=-1):
    """LanguageTool démarré une seule fois ; None s'il est indisponible (-> correcteur Python).
    Si le démarrage est en cours, on attend au plus `attente` secondes (-1 = sans limite)."""
    global tl, tl_essaye
    if not verrou.acquire(timeout=attente):
        print("LanguageTool encore en démarrage -> correcteur Python pour cette fois")
        return None
    try:
        if not tl_essaye:
            tl = demarrer_languagetool()
            tl_essaye = True
    finally:
        verrou.release()
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


@app.on_event("startup")
def demarrer_en_avance():
    # Java met du temps à démarrer sur un petit serveur : on le lance dès le démarrage,
    # en arrière-plan, pour que la 1re correction n'attende pas
    threading.Thread(target=get_outil, daemon=True).start()


@app.get("/")
def read_root():
    return {"message": "Le correcteur fonctionne"}


@app.post("/corriger")
def creer_fichier(file: UploadFile):
    octets = file.file.read()

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
        _, nb = corriger_document(document, data, ignores, get_outil(attente=30), infos)
        statut = "complete" if infos["mode"] in ("local", "public") else "python"
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