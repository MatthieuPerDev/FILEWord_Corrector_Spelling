From d0b3b0422b2f35ec2f70dfdf1d1d621e0e9a6001 Mon Sep 17 00:00:00 2001
From: Claude <noreply@anthropic.com>
Date: Mon, 23 Nov 2026 11:04:09 +0100
Subject: [PATCH] =?UTF-8?q?Correcteur=20utilisable=20avec=20le=20seul=20fi?=
 =?UTF-8?q?chier=20Python=20+=20corrections=20r=C3=A9ellement=20appliqu?=
 =?UTF-8?q?=C3=A9es?=
MIME-Version: 1.0
Content-Type: text/plain; charset=UTF-8
Content-Transfer-Encoding: 8bit

- dictionnaires JSON trouvés quel que soit le dossier de lancement
- tableaux, en-têtes et pieds de page corrigés
- LanguageTool : remplacement à la position exacte (ne remplace plus tous les « la » du paragraphe)
- LanguageTool optionnel : repli automatique sur les dictionnaires
- ligne de commande : chemin complet, nom seul ou fichier glissé

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01KgCsKKehQoCQ5oBMdXnca7
---
 README.md      |  12 ++
 api.py         |  37 ++--
 corrections.py | 517 +++++++++++++++++++++++++++++++++----------------
 ia.py          |   3 +-
 4 files changed, 381 insertions(+), 188 deletions(-)

diff --git a/README.md b/README.md
index e4bbfb4..ab2cff4 100644
--- a/README.md
+++ b/README.md
@@ -27,6 +27,18 @@ Demonstration
 
 If you want to use it, take a look at 
 [this_link..](https://correcteur-word.onrender.com)
+## Use with Python only (no API, no Docker)
+
+```
+pip install python-docx                      # minimum
+pip install language_tool_python wordfreq    # optional (needs Java) for a better correction
+python corrections.py "C:\path\to\file.docx"
+```
+
+The corrected file is saved next to the original as `file_corrige.docx`.
+Options: `--sans-languagetool` (dictionaries only), `--ne-pas-ouvrir` (don't open Word at the end).
+If LanguageTool can't start, the correction still runs with the JSON dictionaries.
+
 ## Author
 
 Matthieu Tchonev Pereira
diff --git a/api.py b/api.py
index 4a68909..6725c44 100644
--- a/api.py
+++ b/api.py
@@ -1,33 +1,27 @@
 # ce fichier permet d'utiliser l'api fastapi
 import io
-import json
-import os
-import language_tool_python
 from docx import Document
 from fastapi import FastAPI, UploadFile, HTTPException
 from fastapi.responses import StreamingResponse
 from fastapi.middleware.cors import CORSMiddleware
-from corrections import corriger_document
+from corrections import corriger_document, charger_json, get_outil as demarrer_languagetool
 from ia import corriger_avec_ia
 
 TAILLE_MAX = 5 * 1024 * 1024
 
-with open("fautes_orthographes/errors.json", "r", encoding="utf-8") as f:
-    data = json.load(f)
-
-with open("fautes_orthographes/ignorer.json", "r", encoding="utf-8") as f:
-    ignores = set(json.load(f))
+data = charger_json("errors.json", {})
+ignores = set(charger_json("ignorer.json", []))
 
 tl = None
+tl_essaye = False
 
 
 def get_outil():
-    global tl
-    if tl is None:
-        if os.getenv("LT_MODE") == "public":
-            tl = language_tool_python.LanguageToolPublicAPI("fr")
-        else:
-            tl = language_tool_python.LanguageTool("fr")
+    """LanguageTool démarré une seule fois ; None s'il est indisponible (-> dictionnaires seuls)."""
+    global tl, tl_essaye
+    if not tl_essaye:
+        tl_essaye = True
+        tl = demarrer_languagetool()
     return tl
 
 
@@ -68,19 +62,14 @@ async def creer_fichier(file: UploadFile):
         print("IA INDISPONIBLE :", e)
         document = Document(io.BytesIO(octets))      # on repart du fichier d'origine
         # plan B : dictionnaire + LanguageTool
-        statut = "complete"
-        try:
-            corriger_document(document, data, ignores, get_outil())
-        except Exception as e:
-            print("LANGUAGETOOL INDISPONIBLE :", e)
-            document = Document(io.BytesIO(octets))
-            corriger_document(document, data, ignores, None)
-            statut = "partielle"
+        outil = get_outil()
+        statut = "complete" if outil is not None else "partielle"
+        corriger_document(document, data, ignores, outil)
 
     sortie = io.BytesIO()
     document.save(sortie)
     sortie.seek(0)
-    nom = file.filename.replace(".docx", "_corrige.docx")
+    nom = (file.filename or "document.docx").replace(".docx", "_corrige.docx")
     return StreamingResponse(  # 7. renvoyer le fichier
         sortie,
         media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
diff --git a/corrections.py b/corrections.py
index 351bd7e..1bbb9a5 100644
--- a/corrections.py
+++ b/corrections.py
@@ -1,58 +1,181 @@
-import subprocess
-import json
-import sys
+"""
+Correcteur d'orthographe pour fichiers Word (.docx) en français.
+
+Ce fichier fonctionne TOUT SEUL (sans l'API, sans Docker, sans l'IA) :
+
+    python corrections.py                      -> demande le nom du fichier
+    python corrections.py "C:\\chemin\\doc.docx"  -> corrige ce fichier
+    python corrections.py doc.docx --sans-languagetool
+
+Minimum requis : pip install python-docx
+Optionnel     : language_tool_python (+ Java) et wordfreq -> correction plus complète.
+Si LanguageTool ne démarre pas, la correction se fait quand même avec les dictionnaires JSON.
+"""
 import bisect
+import json
+import os
 import re
-from docx import Document
-from pathlib import Path  # bibliothèque pour retourner au chemin local du PC
-import language_tool_python
+import subprocess
+import sys
 import unicodedata
-from wordfreq import zipf_frequency
+from pathlib import Path
+
+from docx import Document
 
 try:
-    with open("fautes_orthographes/fautes_generees.json", "r", encoding="utf-8") as f:
-        FAUTES_GENEREES = json.load(f)
-except FileNotFoundError:
-    FAUTES_GENEREES = {}
+    from wordfreq import zipf_frequency
+except ImportError:                       # wordfreq absent -> on fait sans
+    def zipf_frequency(mot, langue):
+        return 0.0
+
+# Dossier du projet : les JSON sont trouvés même si on lance le script d'ailleurs
+DOSSIER = Path(__file__).resolve().parent
+DOSSIER_FAUTES = DOSSIER / "fautes_orthographes"
+
+
+def charger_json(nom, defaut):
+    """Lit un fichier JSON de fautes_orthographes/, ou renvoie `defaut` s'il manque."""
+    chemin = DOSSIER_FAUTES / nom
+    try:
+        with open(chemin, "r", encoding="utf-8") as f:
+            return json.load(f)
+    except FileNotFoundError:
+        print(f"Attention : {chemin} introuvable")
+        return defaut
+
+
+FAUTES_GENEREES = charger_json("fautes_generees.json", {})
+
+
+# ---------------------------------------------------------------------------
+# Parcours du document
+# ---------------------------------------------------------------------------
+
+def paragraphes_du_conteneur(conteneur):
+    """Paragraphes d'un corps / cellule / en-tête, tableaux (même imbriqués) compris."""
+    for p in conteneur.paragraphs:
+        yield p
+    for tableau in conteneur.tables:
+        for ligne in tableau.rows:
+            vues = set()
+            for cellule in ligne.cells:
+                if id(cellule._tc) in vues:          # cellules fusionnées : une seule fois
+                    continue
+                vues.add(id(cellule._tc))
+                yield from paragraphes_du_conteneur(cellule)
+
+
+def tous_les_paragraphes(document):
+    """Tous les paragraphes du document : corps, tableaux, en-têtes et pieds de page."""
+    yield from paragraphes_du_conteneur(document)
+    deja_vus = set()
+    for section in document.sections:
+        for partie in (section.header, section.first_page_header, section.even_page_header,
+                       section.footer, section.first_page_footer, section.even_page_footer):
+            try:
+                if partie.is_linked_to_previous:
+                    continue
+                cle = id(partie._element)
+            except Exception:
+                continue
+            if cle in deja_vus:
+                continue
+            deja_vus.add(cle)
+            yield from paragraphes_du_conteneur(partie)
 
-"""
-Cette fonction permet de corriger un texte avec l'outil LanguageTool,
-sans toucher aux mots de la liste ignores (noms propres, termes techniques)
-"""
-def corriger_texte(texte, outil, mots_ignores):
-    """Corrige un texte et renvoie (texte_corrigé, liste_des_corrections)."""
-    matches = outil.check(texte)
-    corrections = []
-
-    # On part de la FIN pour que les positions des erreurs précédentes restent justes
-    for match in reversed(matches):
-        debut = match.offset                  # début de l'erreur EN COURS
-        fin = debut + match.error_length      # fin de l'erreur EN COURS
-        faute = texte[debut:fin]              # on découpe le TEXTE, pas match
-
-        if faute in mots_ignores:             # nom propre ou terme technique → on ignore
-            continue
-        if not match.replacements:            # pas de suggestion → on laisse tel quel
+
+def texte_runs(paragraphe):
+    """Texte du paragraphe tel qu'il est dans ses runs (c'est ce qu'on peut modifier)."""
+    return "".join(run.text for run in paragraphe.runs)
+
+
+# ---------------------------------------------------------------------------
+# Remplacement dans un paragraphe Word (en gardant la mise en forme)
+# ---------------------------------------------------------------------------
+
+def remplacer_plage(paragraphe, debut, fin, correction):
+    """Remplace les caractères [debut:fin] du texte des runs par `correction`.
+    Seul cet endroit précis est modifié ; le style de chaque run est conservé."""
+    position = 0
+    premier = True
+    for run in paragraphe.runs:
+        texte = run.text
+        r_debut, r_fin = position, position + len(texte)
+        position = r_fin
+        if r_fin <= debut or r_debut >= fin:      # ce run n'est pas concerné
             continue
+        # ce run contient une partie de la plage à remplacer
+        a = max(debut, r_debut) - r_debut
+        b = min(fin, r_fin) - r_debut
+        if premier:
+            run.text = texte[:a] + correction + texte[b:]
+            premier = False
+        else:
+            run.text = texte[:a] + texte[b:]
+        if r_fin >= fin:
+            break
+    return not premier
+
 
-        correction = match.replacements[0]
-        texte = texte[:debut] + correction + texte[fin:]   # avant + correction + après
-        corrections.append((faute, correction, match.message))
+def motif_mot_entier(faute):
+    return re.compile(r"(?<!\w)" + re.escape(faute) + r"(?!\w)")
+
+
+def appliquer_correction(paragraphe, faute, correction):
+    """Remplace chaque `faute` (mot entier) par `correction` dans le paragraphe.
+    Renvoie le nombre de remplacements."""
+    texte = texte_runs(paragraphe)
+    positions = [m.span() for m in motif_mot_entier(faute).finditer(texte)]
+    for debut, fin in reversed(positions):          # depuis la fin : les positions restent justes
+        remplacer_plage(paragraphe, debut, fin, correction)
+    return len(positions)
+
+
+# ---------------------------------------------------------------------------
+# Passe 1 : dictionnaires JSON
+# ---------------------------------------------------------------------------
+
+def avec_majuscules(data):
+    complet = dict(data)
+    for faute, correction in data.items():
+        if faute[:1].islower():
+            complet.setdefault(faute[0].upper() + faute[1:], correction[0].upper() + correction[1:])
+    return complet
 
-    corrections.reverse()   # remettre les corrections dans l'ordre de lecture
-    return texte, corrections
 
+def construire_motif(dictionnaire):
+    """Une seule expression régulière pour toutes les fautes (les plus longues d'abord)."""
+    fautes = sorted(dictionnaire, key=len, reverse=True)
+    if not fautes:
+        return None
+    return re.compile(r"(?<!\w)(?:" + "|".join(re.escape(f) for f in fautes) + r")(?!\w)")
 
 
+def passe_dictionnaire(paragraphes, dictionnaire, motif):
+    nb = 0
+    for p in paragraphes:
+        texte = texte_runs(p)
+        trouves = [(m.start(), m.end(), m.group()) for m in motif.finditer(texte)]
+        for debut, fin, faute in reversed(trouves):
+            if remplacer_plage(p, debut, fin, dictionnaire[faute]):
+                nb += 1
+                print(f"{faute} -> {dictionnaire[faute]} (dictionnaire)")
+    return nb
+
+
+# ---------------------------------------------------------------------------
+# Passe 2 : LanguageTool (optionnel)
+# ---------------------------------------------------------------------------
+
 def distance(a, b):
-    """Nombre de lettres à ajouter, enlever ou changer pour passer de a à b (distance de Levenshtein)."""
+    """Nombre de lettres à ajouter, enlever ou changer pour passer de a à b (Levenshtein)."""
     precedente = list(range(len(b) + 1))
     for i, ca in enumerate(a, 1):
         actuelle = [i]
         for j, cb in enumerate(b, 1):
-            actuelle.append(min(precedente[j] + 1,                 # enlever
-                                actuelle[j - 1] + 1,               # ajouter
-                                precedente[j - 1] + (ca != cb)))   # changer
+            actuelle.append(min(precedente[j] + 1,
+                                actuelle[j - 1] + 1,
+                                precedente[j - 1] + (ca != cb)))
         precedente = actuelle
     return precedente[-1]
 
@@ -60,25 +183,24 @@ def distance(a, b):
 def phonetique(mot):
     """Transforme un mot en une 'clé de son' : simpatik et sympathique donnent la même clé."""
     mot = unicodedata.normalize("NFD", mot.lower())
-    mot = "".join(c for c in mot if unicodedata.category(c) != "Mn")   # enlève les accents
+    mot = "".join(c for c in mot if unicodedata.category(c) != "Mn")
     mot = re.sub(r"[^a-z]", "", mot)
     for avant, apres in [("eau", "o"), ("au", "o"), ("ph", "f"), ("th", "t"), ("qu", "k"),
                          ("c(?=[eiy])", "s"), ("g(?=[eiy])", "j"), ("c", "k"), ("y", "i"),
                          ("ai", "e"), ("ei", "e"), ("e?[ae]n(?=[^aeiou]|$)", "an"), ("z", "s")]:
         mot = re.sub(avant, apres, mot)
-    mot = re.sub(r"(.)\1", r"\1", mot)        # lettres doublées → une seule
-    mot = re.sub(r"[estxd]+$", "", mot)       # lettres muettes à la fin
+    mot = re.sub(r"(.)\1", r"\1", mot)
+    mot = re.sub(r"[estxd]+$", "", mot)
     return mot
 
 
 def choisir_suggestion(faute, suggestions):
     """Choisit la meilleure suggestion de LanguageTool, ou None si aucune n'est assez proche."""
     meilleure, meilleur_score = None, None
-    for suggestion in suggestions[:5]:                         # les 5 premières suffisent
-        # on compare sans apostrophes ni tirets : aujoudui est proche de aujourd'hui
+    for suggestion in suggestions[:5]:
         d = distance(re.sub(r"['’-]", "", faute.lower()), re.sub(r"['’-]", "", suggestion.lower()))
         meme_son = phonetique(faute) == phonetique(suggestion) and phonetique(faute) != ""
-        if d > max(2, len(faute) // 4) and not meme_son:     # trop éloignée → on ne la prend pas
+        if d > max(2, len(faute) // 4) and not meme_son:
             continue
         score = zipf_frequency(suggestion, "fr") - d + (3 if meme_son else 0)
         if meilleur_score is None or score > meilleur_score:
@@ -86,138 +208,207 @@ def choisir_suggestion(faute, suggestions):
     return meilleure
 
 
-def motif_mot_entier(faute):
-    return re.compile(r"(?<!\w)" + re.escape(faute) + r"(?!\w)")
-
-
-def avec_majuscules(data):
-    complet = dict(data)
-    for faute, correction in data.items():
-        if faute[:1].islower():
-            complet.setdefault(faute[0].upper() + faute[1:], correction[0].upper() + correction[1:])
-    return complet
-
+def get_outil(mode=None):
+    """Démarre LanguageTool. Renvoie None s'il est indisponible (pas installé, pas de Java...)."""
+    mode = mode or os.getenv("LT_MODE", "local")
+    try:
+        import language_tool_python
+    except ImportError:
+        print("language_tool_python n'est pas installé -> correction avec les dictionnaires seulement")
+        return None
+    essais = ["public"] if mode == "public" else ["local", "public"]
+    for essai in essais:
+        try:
+            if essai == "local":
+                outil = language_tool_python.LanguageTool("fr")
+            else:
+                outil = language_tool_python.LanguageToolPublicAPI("fr")
+            outil.check("Test.")                     # vérifie qu'il répond vraiment
+            print(f"LanguageTool prêt ({essai})")
+            return outil
+        except Exception as e:
+            print(f"LanguageTool {essai} indisponible : {e}")
+    print("-> correction avec les dictionnaires seulement")
+    return None
+
+
+def paquets_de_paragraphes(paragraphes, taille_max=15000):
+    """Regroupe les paragraphes en paquets de texte (l'API publique limite la taille)."""
+    paquet, taille = [], 0
+    for p in paragraphes:
+        n = len(texte_runs(p)) + 1
+        if paquet and taille + n > taille_max:
+            yield paquet
+            paquet, taille = [], 0
+        paquet.append(p)
+        taille += n
+    if paquet:
+        yield paquet
+
+
+def passe_languagetool(paragraphes, ignores, tl):
+    nb = 0
+    for tour in range(2):     # 2 tours : certaines fautes n'apparaissent qu'une fois les premières corrigées
+        corrige_ce_tour = 0
+        for paquet in paquets_de_paragraphes(paragraphes):
+            textes = [texte_runs(p) for p in paquet]
+            texte_complet = "\n".join(textes)
+            matches = tl.check(texte_complet)
+
+            debuts, position = [], 0
+            for texte in textes:
+                debuts.append(position)
+                position += len(texte) + 1
+
+            # on regroupe les corrections par paragraphe, puis on les applique depuis la fin
+            a_faire = {}
+            for match in matches:
+                faute = texte_complet[match.offset: match.offset + match.error_length]
+                if not faute.strip() or "\n" in faute or not match.replacements:
+                    continue
+                if faute in ignores or any(mot in faute for mot in ignores):
+                    continue
+                correction = choisir_suggestion(faute, match.replacements)
+                if correction is None:
+                    print(f"{faute} : laissé tel quel ({match.replacements[:3]})")
+                    continue
+                i = bisect.bisect_right(debuts, match.offset) - 1
+                debut = match.offset - debuts[i]
+                a_faire.setdefault(i, []).append((debut, debut + len(faute), faute, correction, match.message))
+
+            for i, liste in a_faire.items():
+                fin_precedente = None
+                for debut, fin, faute, correction, message in sorted(liste, reverse=True):
+                    if fin_precedente is not None and fin > fin_precedente:   # chevauchement -> on saute
+                        continue
+                    if remplacer_plage(paquet[i], debut, fin, correction):
+                        nb += 1
+                        corrige_ce_tour += 1
+                        fin_precedente = debut
+                        print(f"{faute} -> {correction} ({message})")
+        if corrige_ce_tour == 0:
+            break
+    return nb
 
-def appliquer_correction(paragraphe, faute, correction):
-    """Remplace `faute` par `correction` dans un paragraphe Word en gardant la mise en forme.
-    Renvoie True si la correction a été appliquée, False sinon."""
-    # Sortie anticipée : la faute n'est pas dans ce paragraphe
-    if faute not in paragraphe.text:
-        return False
-    motif = motif_mot_entier(faute)
-    if not motif.search(paragraphe.text):
-        return False
-
-    # Étape 1 : correction run par run (garde le style de chaque morceau)
-    corrige = False
-    for run in paragraphe.runs:  # chaque morceau du paragraphe
-        if motif.search(run.text):
-            run.text = motif.sub(lambda m: correction, run.text)
-            corrige = True
-
-    # Étape 2 : la faute est coupée entre plusieurs runs
-    if not corrige:
-        nouveau_texte = motif.sub(lambda m: correction, paragraphe.text)
-        runs_texte = [run for run in paragraphe.runs if run.text]  # runs avec du texte (pas les images)
-        runs_texte[0].text = nouveau_texte                         # tout le texte corrigé dans le 1er run
-        for run in runs_texte[1:]:
-            run.text = ""                                          # on vide les autres
-
-    return True
 
+# ---------------------------------------------------------------------------
+# Correction complète
+# ---------------------------------------------------------------------------
 
-def main():
-    file_name = input("Quel est le nom de votre fichier à corriger (sans extension) : \n")
+def corriger_document(document, data=None, ignores=None, tl=None):
+    """Corrige le document (modifié sur place). Renvoie (nombre_de_mots, nombre_de_corrections).
+    tl = outil LanguageTool, ou None pour corriger seulement avec les dictionnaires."""
+    if data is None:
+        data = charger_json("errors.json", {})
+    if ignores is None:
+        ignores = set(charger_json("ignorer.json", []))
 
-    home = Path.home()           # root path
-    desktop = home / "Desktop"   # bureau
-    resultats = list(desktop.rglob(file_name + ".docx"))  # liste de tous les fichiers trouvés portant ce nom
-    if len(resultats) == 0:
-        print("Impossible de trouver le fichier correspondant sur votre machine \n")
-        print("Essayez une autre syntaxe")
-        sys.exit()
-    if len(resultats) > 1:
-        print("Faites attention, il y a un autre fichier qui porte le même nom")
+    dictionnaire = {f: c for f, c in {**FAUTES_GENEREES, **avec_majuscules(data)}.items()
+                    if f not in ignores and f != c}
+    paragraphes = [p for p in tous_les_paragraphes(document) if texte_runs(p).strip()]
+    cmpt = sum(len(texte_runs(p).split()) for p in paragraphes)
 
-    path_file = resultats[0]
-    document = Document(path_file)  # sert à parcourir les paragraphes
+    nb_corrections = 0
+    motif = construire_motif(dictionnaire)
+    if motif:
+        nb_corrections += passe_dictionnaire(paragraphes, dictionnaire, motif)
 
-    # ouverture des JSON
-    with open("fautes_orthographes/errors.json", "r", encoding='utf-8') as f1:
-        data = json.load(f1)            # fautes et corrections personnelles
+    if tl is not None:
+        try:
+            nb_corrections += passe_languagetool(paragraphes, ignores, tl)
+        except Exception as e:     # LanguageTool plante en cours de route -> on garde ce qui est fait
+            print("LanguageTool a échoué :", e)
 
-    with open("fautes_orthographes/ignorer.json", "r", encoding='utf-8') as f2:
-        ignores = set(json.load(f2))    # mots à ne pas corriger
+    return cmpt, nb_corrections
 
-    tl = language_tool_python.LanguageTool('fr')  # créé UNE SEULE FOIS, avant la boucle
 
-    cmpt, nb_corrections = corriger_document(document, data, ignores, tl)  # toute la correction
+# ---------------------------------------------------------------------------
+# Utilisation en ligne de commande
+# ---------------------------------------------------------------------------
+
+def trouver_fichier(saisie):
+    """Accepte un chemin complet, un nom avec ou sans .docx, ou un fichier glissé dans le terminal."""
+    saisie = saisie.strip().strip('"').strip("'").strip()
+    if not saisie:
+        return None
+    chemin = Path(saisie).expanduser()
+    if chemin.suffix.lower() != ".docx":
+        chemin = chemin.with_name(chemin.name + ".docx")
+    if chemin.is_file():
+        return chemin
+
+    nom = chemin.name
+    home = Path.home()
+    trouves = []
+    for dossier in [Path.cwd(), home / "Desktop", home / "Bureau", home / "OneDrive" / "Desktop",
+                    home / "OneDrive" / "Bureau", home / "Documents", home / "Downloads",
+                    home / "Téléchargements"]:
+        if dossier.is_dir():
+            try:
+                trouves += [f for f in dossier.rglob(nom) if not f.name.startswith("~$")]
+            except OSError:
+                pass
+    if not trouves:
+        return None
+    if len(trouves) > 1:
+        print("Attention, plusieurs fichiers portent ce nom, j'utilise :", trouves[0])
+    return trouves[0]
+
+
+def ouvrir_fichier(chemin):
+    try:
+        if sys.platform.startswith("win"):
+            os.startfile(chemin)
+        elif sys.platform == "darwin":
+            subprocess.run(["open", str(chemin)])
+        else:
+            subprocess.run(["xdg-open", str(chemin)])
+    except Exception:
+        pass
 
-    tl.close()  # on ferme l'outil
 
-    print(cmpt)
+def main():
+    args = [a for a in sys.argv[1:] if not a.startswith("--")]
+    sans_lt = "--sans-languagetool" in sys.argv
+    ne_pas_ouvrir = "--ne-pas-ouvrir" in sys.argv
+
+    saisie = args[0] if args else input("Nom ou chemin du fichier Word à corriger : \n")
+    chemin = trouver_fichier(saisie)
+    if chemin is None:
+        print("Impossible de trouver ce fichier. Donnez le chemin complet, par exemple :")
+        print(r'  python corrections.py "C:\Users\moi\Desktop\mon_fichier.docx"')
+        sys.exit(1)
+
+    print("Fichier :", chemin)
+    try:
+        document = Document(str(chemin))
+    except Exception as e:
+        print("Ce fichier n'est pas un document Word valide (.docx) :", e)
+        sys.exit(1)
+
+    tl = None if sans_lt else get_outil()
+    try:
+        cmpt, nb_corrections = corriger_document(document, tl=tl)
+    finally:
+        if tl is not None:
+            try:
+                tl.close()
+            except Exception:
+                pass
+
+    print("Nombre de mots :", cmpt)
     print("Nombre de corrections :", nb_corrections)
 
-    chemin_corrige = path_file.with_stem(path_file.stem + "_corrige")
-    document.save(chemin_corrige)
-    subprocess.run(["cmd", "/c", "start", "winword", str(chemin_corrige)])  # ouvre le fichier avec Word
-
-
-def corriger_document(document, data, ignores, tl):
-    cmpt = 0  # compteur de mots
-    nb_corrections = 0  # compteur de corrections
-
-    data = {**FAUTES_GENEREES, **avec_majuscules(data)}
-    paragraphes = [p for p in document.paragraphs if p.text.strip()]  # paragraphe vide → on passe au suivant
-
-    # Passe 1 : on cherche dans le fichier avec les fautes et les corrections
-    for paragraphe in paragraphes:
-        texte = paragraphe.text
-        cmpt += len(texte.split())
-        for faute, correction in data.items():
-            if faute in ignores or faute not in texte:
-                continue
-            if appliquer_correction(paragraphe, faute, correction):
-                nb_corrections += 1
-                texte = paragraphe.text
-
-    if tl is None:
-        return cmpt, nb_corrections
-
-    # Passe 2 : LanguageTool, 2 tours : certaines fautes n'apparaissent qu'une fois les premières corrigées
-    for tour in range(2):
-        textes = [p.text for p in paragraphes]
-        texte_complet = "\n".join(textes)
-        matches = tl.check(texte_complet)
-
-        debuts = []
-        position = 0
-        for texte in textes:
-            debuts.append(position)
-            position += len(texte) + 1
-
-        corrige_ce_tour = 0
-        for match in matches:
-            faute = texte_complet[match.offset : match.offset + match.error_length]
-            if faute in ignores or not match.replacements:
-                continue
-
-            correction = choisir_suggestion(faute, match.replacements)
-            if correction is None:      # aucune suggestion assez proche → mieux vaut laisser la faute
-                print(f"{faute} : laissé tel quel ({match.replacements[:3]})")
-                continue
-
-            i = bisect.bisect_right(debuts, match.offset) - 1
-            if appliquer_correction(paragraphes[i], faute, correction):
-                nb_corrections += 1
-                corrige_ce_tour += 1
-                print(f"{faute} -> {correction} ({match.message})")
-
-        if corrige_ce_tour == 0:        # rien de nouveau → pas besoin d'un 2e tour
-            break
-
-    return cmpt, nb_corrections
+    chemin_corrige = chemin.with_name(chemin.stem + "_corrige.docx")
+    try:
+        document.save(str(chemin_corrige))
+    except PermissionError:
+        print(f"Impossible d'enregistrer {chemin_corrige} : fermez-le dans Word puis relancez.")
+        sys.exit(1)
+    print("Fichier corrigé enregistré :", chemin_corrige)
+    if not ne_pas_ouvrir:
+        ouvrir_fichier(chemin_corrige)
 
 
 if __name__ == "__main__":
-    main()
\ No newline at end of file
+    main()
diff --git a/ia.py b/ia.py
index 7fba2ff..d966420 100644
--- a/ia.py
+++ b/ia.py
@@ -105,8 +105,9 @@ def remplacer_texte(paragraphe, nouveau):
 
 def corriger_avec_ia(document, ignores):
     """Corrige tout le document avec l'IA. Lève une erreur si l'IA ne répond pas (→ plan B)."""
+    from corrections import tous_les_paragraphes   # corps + tableaux + en-têtes/pieds de page
     paragraphes = []
-    for p in document.paragraphs:
+    for p in tous_les_paragraphes(document):
         texte = "".join(run.text for run in p.runs)
         if texte.strip() and texte == p.text:        # on saute les paragraphes avec des liens (texte hors runs)
             paragraphes.append(p)
-- 
2.43.0