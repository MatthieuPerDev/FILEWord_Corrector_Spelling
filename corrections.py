"""
Correcteur d'orthographe pour fichiers Word (.docx) en français.

Ce fichier fonctionne TOUT SEUL (sans l'API, sans Docker, sans l'IA) :

    python corrections.py                      -> demande le nom du fichier
    python corrections.py "C:\\chemin\\doc.docx"  -> corrige ce fichier
    python corrections.py doc.docx --sans-languagetool

Minimum requis : pip install python-docx
Optionnel     : language_tool_python (+ Java) et wordfreq -> correction plus complète.
Si LanguageTool ne démarre pas, la correction se fait quand même avec les dictionnaires JSON.
"""
import bisect
import json
import os
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

from docx import Document

try:
    from wordfreq import zipf_frequency
    WORDFREQ_OK = True
except ImportError:                       # wordfreq absent -> on fait sans
    WORDFREQ_OK = False

    def zipf_frequency(mot, langue):
        return 0.0

# Dossier du projet : les JSON sont trouvés même si on lance le script d'ailleurs
DOSSIER = Path(__file__).resolve().parent
DOSSIER_FAUTES = DOSSIER / "fautes_orthographes"


def charger_json(nom, defaut):
    """Lit un fichier JSON de fautes_orthographes/, ou renvoie `defaut` s'il manque."""
    chemin = DOSSIER_FAUTES / nom
    try:
        with open(chemin, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Attention : {chemin} introuvable")
        return defaut


FAUTES_GENEREES = charger_json("fautes_generees.json", {})


# ---------------------------------------------------------------------------
# Parcours du document
# ---------------------------------------------------------------------------

def paragraphes_du_conteneur(conteneur):
    """Paragraphes d'un corps / cellule / en-tête, tableaux (même imbriqués) compris."""
    for p in conteneur.paragraphs:
        yield p
    for tableau in conteneur.tables:
        for ligne in tableau.rows:
            vues = set()
            for cellule in ligne.cells:
                if id(cellule._tc) in vues:          # cellules fusionnées : une seule fois
                    continue
                vues.add(id(cellule._tc))
                yield from paragraphes_du_conteneur(cellule)


def tous_les_paragraphes(document):
    """Tous les paragraphes du document : corps, tableaux, en-têtes et pieds de page."""
    yield from paragraphes_du_conteneur(document)
    deja_vus = set()
    for section in document.sections:
        for partie in (section.header, section.first_page_header, section.even_page_header,
                       section.footer, section.first_page_footer, section.even_page_footer):
            try:
                if partie.is_linked_to_previous:
                    continue
                cle = id(partie._element)
            except Exception:
                continue
            if cle in deja_vus:
                continue
            deja_vus.add(cle)
            yield from paragraphes_du_conteneur(partie)


def texte_runs(paragraphe):
    """Texte du paragraphe tel qu'il est dans ses runs (c'est ce qu'on peut modifier)."""
    return "".join(run.text for run in paragraphe.runs)


# ---------------------------------------------------------------------------
# Remplacement dans un paragraphe Word (en gardant la mise en forme)
# ---------------------------------------------------------------------------

def remplacer_plage(paragraphe, debut, fin, correction):
    """Remplace les caractères [debut:fin] du texte des runs par `correction`.
    Seul cet endroit précis est modifié ; le style de chaque run est conservé."""
    position = 0
    premier = True
    for run in paragraphe.runs:
        texte = run.text
        r_debut, r_fin = position, position + len(texte)
        position = r_fin
        if r_fin <= debut or r_debut >= fin:      # ce run n'est pas concerné
            continue
        # ce run contient une partie de la plage à remplacer
        a = max(debut, r_debut) - r_debut
        b = min(fin, r_fin) - r_debut
        if premier:
            run.text = texte[:a] + correction + texte[b:]
            premier = False
        else:
            run.text = texte[:a] + texte[b:]
        if r_fin >= fin:
            break
    return not premier


def motif_mot_entier(faute):
    return re.compile(r"(?<!\w)" + re.escape(faute) + r"(?!\w)")


def appliquer_correction(paragraphe, faute, correction):
    """Remplace chaque `faute` (mot entier) par `correction` dans le paragraphe.
    Renvoie le nombre de remplacements."""
    texte = texte_runs(paragraphe)
    positions = [m.span() for m in motif_mot_entier(faute).finditer(texte)]
    for debut, fin in reversed(positions):          # depuis la fin : les positions restent justes
        remplacer_plage(paragraphe, debut, fin, correction)
    return len(positions)


# ---------------------------------------------------------------------------
# Passe 1 : dictionnaires JSON
# ---------------------------------------------------------------------------

def avec_majuscules(data):
    complet = dict(data)
    for faute, correction in data.items():
        if faute[:1].islower():
            complet.setdefault(faute[0].upper() + faute[1:], correction[0].upper() + correction[1:])
    return complet


def construire_motif(dictionnaire):
    """Une seule expression régulière pour toutes les fautes (les plus longues d'abord)."""
    fautes = sorted(dictionnaire, key=len, reverse=True)
    if not fautes:
        return None
    return re.compile(r"(?<!\w)(?:" + "|".join(re.escape(f) for f in fautes) + r")(?!\w)")


def passe_dictionnaire(paragraphes, dictionnaire, motif):
    nb = 0
    for p in paragraphes:
        texte = texte_runs(p)
        trouves = [(m.start(), m.end(), m.group()) for m in motif.finditer(texte)]
        for debut, fin, faute in reversed(trouves):
            if remplacer_plage(p, debut, fin, dictionnaire[faute]):
                nb += 1
                print(f"{faute} -> {dictionnaire[faute]} (dictionnaire)")
    return nb


# ---------------------------------------------------------------------------
# Passe 2 : LanguageTool (optionnel)
# ---------------------------------------------------------------------------

def distance(a, b):
    """Nombre de lettres à ajouter, enlever ou changer pour passer de a à b (Levenshtein)."""
    precedente = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        actuelle = [i]
        for j, cb in enumerate(b, 1):
            actuelle.append(min(precedente[j] + 1,
                                actuelle[j - 1] + 1,
                                precedente[j - 1] + (ca != cb)))
        precedente = actuelle
    return precedente[-1]


def phonetique(mot):
    """Transforme un mot en une 'clé de son' : simpatik et sympathique donnent la même clé."""
    mot = unicodedata.normalize("NFD", mot.lower())
    mot = "".join(c for c in mot if unicodedata.category(c) != "Mn")
    mot = re.sub(r"[^a-z]", "", mot)
    for avant, apres in [("eau", "o"), ("au", "o"), ("ph", "f"), ("th", "t"), ("qu", "k"),
                         ("c(?=[eiy])", "s"), ("g(?=[eiy])", "j"), ("c", "k"), ("y", "i"),
                         ("ai", "e"), ("ei", "e"), ("e?[ae]n(?=[^aeiou]|$)", "an"), ("z", "s")]:
        mot = re.sub(avant, apres, mot)
    mot = re.sub(r"(.)\1", r"\1", mot)
    mot = re.sub(r"[estxd]+$", "", mot)
    return mot


def choisir_suggestion(faute, suggestions):
    """Choisit la meilleure suggestion de LanguageTool, ou None si aucune n'est assez proche."""
    meilleure, meilleur_score = None, None
    for suggestion in suggestions[:5]:
        d = distance(re.sub(r"['’-]", "", faute.lower()), re.sub(r"['’-]", "", suggestion.lower()))
        meme_son = phonetique(faute) == phonetique(suggestion) and phonetique(faute) != ""
        if d > max(2, len(faute) // 4) and not meme_son:
            continue
        score = zipf_frequency(suggestion, "fr") - d + (3 if meme_son else 0)
        if meilleur_score is None or score > meilleur_score:
            meilleure, meilleur_score = suggestion, score
    return meilleure


def creer_local(language_tool_python):
    """LanguageTool local (Java). La bibliothèque n'attend que 15 s que Java démarre :
    sur un petit serveur (Render gratuit) c'est trop court, on attend jusqu'à LT_ATTENTE secondes."""
    attente = int(os.getenv("LT_ATTENTE", "300"))

    class LanguageToolPatient(language_tool_python.LanguageTool):
        def _wait_for_server_ready(self, timeout=15):
            return super()._wait_for_server_ready(timeout=max(timeout, attente))

    return LanguageToolPatient("fr")


def get_outil(mode=None):
    """Démarre LanguageTool. Renvoie None s'il est indisponible (pas installé, pas de Java...)."""
    mode = mode or os.getenv("LT_MODE", "local")
    try:
        import language_tool_python
    except ImportError:
        print("language_tool_python n'est pas installé -> correction avec les dictionnaires seulement")
        return None
    essais = ["public"] if mode == "public" else ["local", "public"]
    for essai in essais:
        try:
            if essai == "local":
                outil = creer_local(language_tool_python)
            else:
                outil = language_tool_python.LanguageToolPublicAPI("fr")
            outil.check("Test.")                     # vérifie qu'il répond vraiment
            print(f"LanguageTool prêt ({essai})")
            return outil
        except Exception as e:
            print(f"LanguageTool {essai} indisponible : {e}")
    print("-> correction avec les dictionnaires seulement")
    return None


_outil_public = None


def outil_public():
    """LanguageTool en ligne (pas besoin de Java), créé une seule fois."""
    global _outil_public
    if _outil_public is None:
        import language_tool_python
        _outil_public = language_tool_python.LanguageToolPublicAPI("fr")
    return _outil_public


def est_public(tl):
    return type(tl).__name__ == "LanguageToolPublicAPI"


def paquets_de_paragraphes(paragraphes, taille_max=15000):
    """Regroupe les paragraphes en paquets de texte (l'API publique limite la taille)."""
    paquet, taille = [], 0
    for p in paragraphes:
        n = len(texte_runs(p)) + 1
        if paquet and taille + n > taille_max:
            yield paquet
            paquet, taille = [], 0
        paquet.append(p)
        taille += n
    if paquet:
        yield paquet


def passe_languagetool(paragraphes, ignores, tl):
    nb = 0
    for tour in range(2):     # 2 tours : certaines fautes n'apparaissent qu'une fois les premières corrigées
        corrige_ce_tour = 0
        for paquet in paquets_de_paragraphes(paragraphes):
            textes = [texte_runs(p) for p in paquet]
            texte_complet = "\n".join(textes)
            matches = tl.check(texte_complet)

            debuts, position = [], 0
            for texte in textes:
                debuts.append(position)
                position += len(texte) + 1

            # on regroupe les corrections par paragraphe, puis on les applique depuis la fin
            a_faire = {}
            for match in matches:
                faute = texte_complet[match.offset: match.offset + match.error_length]
                if not faute.strip() or "\n" in faute or not match.replacements:
                    continue
                if faute in ignores or any(mot in faute for mot in ignores):
                    continue
                correction = choisir_suggestion(faute, match.replacements)
                if correction is None:
                    print(f"{faute} : laissé tel quel ({match.replacements[:3]})")
                    continue
                i = bisect.bisect_right(debuts, match.offset) - 1
                debut = match.offset - debuts[i]
                a_faire.setdefault(i, []).append((debut, debut + len(faute), faute, correction, match.message))

            for i, liste in a_faire.items():
                fin_precedente = None
                for debut, fin, faute, correction, message in sorted(liste, reverse=True):
                    if fin_precedente is not None and fin > fin_precedente:   # chevauchement -> on saute
                        continue
                    if remplacer_plage(paquet[i], debut, fin, correction):
                        nb += 1
                        corrige_ce_tour += 1
                        fin_precedente = debut
                        print(f"{faute} -> {correction} ({message})")
        if corrige_ce_tour == 0:
            break
    return nb


# ---------------------------------------------------------------------------
# Passe 3 : correcteur 100 % Python (sans Java, sans Internet)
# Utilisé quand LanguageTool est indisponible. Règles prudentes : en cas de doute, on ne touche pas.
# ---------------------------------------------------------------------------

LETTRES = "A-Za-zÀ-ÖØ-öø-ÿŒœÆæ"
RE_MOT = re.compile(f"[{LETTRES}]+")
ALPHABET = "abcdefghijklmnopqrstuvwxyzàâäçéèêëîïôöùûüÿœ"

# je/tu + verbe écrit à la 3e personne : "je fait" -> "je fais"
VERBES_JE_TU = {
    "fait": "fais", "doit": "dois", "peut": "peux", "veut": "veux", "sait": "sais", "dit": "dis",
    "voit": "vois", "met": "mets", "prend": "prends", "vient": "viens", "tient": "tiens",
    "part": "pars", "sort": "sors", "dort": "dors", "croit": "crois", "écrit": "écris",
    "lit": "lis", "connait": "connais", "connaît": "connais", "comprend": "comprends",
    "apprend": "apprends", "attend": "attends", "entend": "entends", "répond": "réponds",
    "perd": "perds", "vend": "vends", "vit": "vis", "finit": "finis", "choisit": "choisis",
    "réussit": "réussis", "boit": "bois", "court": "cours", "sent": "sens", "sert": "sers",
    "reçoit": "reçois", "devient": "deviens", "revient": "reviens", "obtient": "obtiens",
    "permet": "permets", "promet": "promets", "remet": "remets",
}
VERBES_TU = {"a": "as", "va": "vas", "est": "es"}
VERBES_JE = {"va": "vais"}
PRONOMS = r"(?:(?:ne|me|te|le|la|les|lui|leur|en|y)\s+|(?:n|m|t|l)['’]\s*)*"
RE_JE_TU = re.compile(rf"(?<![\w'’])(je|tu)(\s+{PRONOMS})([{LETTRES}]+)(?![\w'’-])", re.IGNORECASE)

# "tu pense" -> "tu penses" : ces mots finissent par "e" mais ne sont pas des verbes
PAS_DES_VERBES = {"le", "la", "me", "te", "se", "ne", "que", "comme", "même", "une", "elle", "cette",
                  "notre", "votre", "quelle", "quelque", "entre", "contre", "presque", "encore",
                  "toute", "chaque", "rare", "simple", "seule", "juste", "libre", "possible"}

# déterminants toujours pluriels : "des faute" -> "des fautes"
DETERMINANTS_PLURIEL = {"des", "mes", "tes", "ses", "nos", "vos", "ces", "plusieurs", "quelques", "leurs"}
INVARIABLES = {"le", "la", "les", "un", "une", "de", "du", "des", "en", "et", "ou", "que", "qui", "pour",
               "par", "sur", "sous", "avec", "sans", "dans", "plus", "moins", "très", "trop", "bien",
               "mal", "mille", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", "dix",
               "vingt", "cent", "ce", "ma", "ta", "sa", "mon", "ton", "son", "au", "aux", "ne", "se", "lors"}
NE_PAS_CORRIGER = {"ores"}


def meme_casse(modele, mot):
    """Garde la majuscule du mot d'origine."""
    if modele[:1].isupper():
        return mot[:1].upper() + mot[1:]
    return mot


def suggestion_proche(mot):
    """Mot très proche (1 lettre de différence) et au moins 300 fois plus fréquent, sinon None."""
    m = mot.lower()
    freq = zipf_frequency(m, "fr")
    if freq >= 3:                                  # mot courant -> on ne touche pas
        return None
    if zipf_frequency(m, "en") >= 3:               # mot anglais courant -> on ne touche pas
        return None
    candidats = set()
    for i in range(len(m) + 1):
        a, b = m[:i], m[i:]
        if b:
            candidats.add(a + b[1:])                                   # lettre en trop
            candidats.update(a + c + b[1:] for c in ALPHABET)          # mauvaise lettre
        if len(b) > 1:
            candidats.add(a + b[1] + b[0] + b[2:])                     # lettres inversées
        candidats.update(a + c + b for c in ALPHABET)                  # lettre oubliée
    candidats.discard(m)
    meilleur, meilleure_freq = None, 0
    for c in candidats:
        f = zipf_frequency(c, "fr")
        if f > meilleure_freq:
            meilleur, meilleure_freq = c, f
    if meilleur and meilleure_freq >= 3 and meilleure_freq - freq >= 2.5:
        return meilleur
    return None


def corriger_phrase_python(texte, ignores):
    """Renvoie la liste des corrections (debut, fin, faute, correction) pour un texte."""
    corrections = []

    # a) je/tu + verbe
    for m in RE_JE_TU.finditer(texte):
        sujet, verbe = m.group(1).lower(), m.group(3)
        v = verbe.lower()
        bon = VERBES_JE_TU.get(v) or (VERBES_TU.get(v) if sujet == "tu" else VERBES_JE.get(v))
        if bon is None and sujet == "tu" and v.endswith("e") and len(v) >= 3 and v not in PAS_DES_VERBES \
                and zipf_frequency(v + "s", "fr") >= 2:
            bon = v + "s"                          # tu pense -> tu penses
        if bon:
            corrections.append((m.start(3), m.end(3), verbe, meme_casse(verbe, bon)))

    # b) déterminant pluriel + nom au singulier
    mots = list(RE_MOT.finditer(texte))
    for avant, mot in zip(mots, mots[1:]):
        if avant.group().lower() not in DETERMINANTS_PLURIEL:
            continue
        if texte[avant.end():mot.start()].strip():     # ponctuation entre les deux -> on ne touche pas
            continue
        w = mot.group()
        if not w.islower() or len(w) < 3 or w[-1] in "sxz" or w in INVARIABLES or w in ignores:
            continue
        if zipf_frequency(w, "fr") >= 1 and zipf_frequency(w + "s", "fr") >= 2:
            corrections.append((mot.start(), mot.end(), w, w + "s"))

    # c) fautes de frappe : porcc -> porc, longtemp -> longtemps
    if WORDFREQ_OK:
        debut_phrase = True
        for m in RE_MOT.finditer(texte):
            w = m.group()
            precedent = texte[:m.start()].rstrip()
            debut_phrase = not precedent or precedent[-1] in ".!?:\n"
            if len(w) <= 3 or w in ignores or w.lower() in NE_PAS_CORRIGER:
                continue
            if any(c.isupper() for c in w[1:]):                 # sigles, DevOps...
                continue
            if w[0].isupper() and not debut_phrase:             # nom propre
                continue
            bon = suggestion_proche(w)
            if bon:
                corrections.append((m.start(), m.end(), w, meme_casse(w, bon)))

    # on garde une seule correction par endroit (pas de chevauchement)
    corrections.sort()
    gardees, fin_precedente = [], -1
    for c in corrections:
        if c[0] >= fin_precedente:
            gardees.append(c)
            fin_precedente = c[1]
    return gardees


def passe_python(paragraphes, ignores):
    nb = 0
    for p in paragraphes:
        texte = texte_runs(p)
        for debut, fin, faute, correction in reversed(corriger_phrase_python(texte, ignores)):
            if faute != correction and remplacer_plage(p, debut, fin, correction):
                nb += 1
                print(f"{faute} -> {correction} (correcteur Python)")
    return nb


# ---------------------------------------------------------------------------
# Correction complète
# ---------------------------------------------------------------------------

def corriger_document(document, data=None, ignores=None, tl=None, infos=None):
    """Corrige le document (modifié sur place). Renvoie (nombre_de_mots, nombre_de_corrections).
    tl = outil LanguageTool, ou None pour corriger seulement avec les dictionnaires.
    infos (dict, optionnel) reçoit "mode" : "local", "public" ou "python"."""
    if data is None:
        data = charger_json("errors.json", {})
    if ignores is None:
        ignores = set(charger_json("ignorer.json", []))

    dictionnaire = {f: c for f, c in {**FAUTES_GENEREES, **avec_majuscules(data)}.items()
                    if f not in ignores and f != c}
    paragraphes = [p for p in tous_les_paragraphes(document) if texte_runs(p).strip()]
    cmpt = sum(len(texte_runs(p).split()) for p in paragraphes)

    nb_corrections = 0
    motif = construire_motif(dictionnaire)
    if motif:
        nb_corrections += passe_dictionnaire(paragraphes, dictionnaire, motif)

    mode = "dictionnaires"
    if tl is not None:
        try:
            nb_corrections += passe_languagetool(paragraphes, ignores, tl)
            mode = "public" if est_public(tl) else "local"
        except Exception as e:     # LanguageTool plante en cours de route (ex : manque de mémoire)
            print("LanguageTool a échoué :", e)
            if not est_public(tl):
                try:               # plan B : LanguageTool en ligne
                    nb_corrections += passe_languagetool(paragraphes, ignores, outil_public())
                    mode = "public"
                    print("-> LanguageTool public utilisé à la place")
                except Exception as e2:
                    print("LanguageTool public a échoué aussi :", e2)

    if mode == "dictionnaires":    # pas de LanguageTool -> correcteur 100 % Python
        nb_corrections += passe_python(paragraphes, ignores)
        mode = "python"

    if infos is not None:
        infos["mode"] = mode
    return cmpt, nb_corrections


# ---------------------------------------------------------------------------
# Utilisation en ligne de commande
# ---------------------------------------------------------------------------

def trouver_fichier(saisie):
    """Accepte un chemin complet, un nom avec ou sans .docx, ou un fichier glissé dans le terminal."""
    saisie = saisie.strip().strip('"').strip("'").strip()
    if not saisie:
        return None
    chemin = Path(saisie).expanduser()
    if chemin.suffix.lower() != ".docx":
        chemin = chemin.with_name(chemin.name + ".docx")
    if chemin.is_file():
        return chemin

    nom = chemin.name
    home = Path.home()
    trouves = []
    for dossier in [Path.cwd(), home / "Desktop", home / "Bureau", home / "OneDrive" / "Desktop",
                    home / "OneDrive" / "Bureau", home / "Documents", home / "Downloads",
                    home / "Téléchargements"]:
        if dossier.is_dir():
            try:
                trouves += [f for f in dossier.rglob(nom) if not f.name.startswith("~$")]
            except OSError:
                pass
    if not trouves:
        return None
    if len(trouves) > 1:
        print("Attention, plusieurs fichiers portent ce nom, j'utilise :", trouves[0])
    return trouves[0]


def ouvrir_fichier(chemin):
    try:
        if sys.platform.startswith("win"):
            os.startfile(chemin)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(chemin)])
        else:
            subprocess.run(["xdg-open", str(chemin)])
    except Exception:
        pass


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    sans_lt = "--sans-languagetool" in sys.argv
    ne_pas_ouvrir = "--ne-pas-ouvrir" in sys.argv

    saisie = args[0] if args else input("Nom ou chemin du fichier Word à corriger : \n")
    chemin = trouver_fichier(saisie)
    if chemin is None:
        print("Impossible de trouver ce fichier. Donnez le chemin complet, par exemple :")
        print(r'  python corrections.py "C:\Users\moi\Desktop\mon_fichier.docx"')
        sys.exit(1)

    print("Fichier :", chemin)
    try:
        document = Document(str(chemin))
    except Exception as e:
        print("Ce fichier n'est pas un document Word valide (.docx) :", e)
        sys.exit(1)

    tl = None if sans_lt else get_outil()
    try:
        cmpt, nb_corrections = corriger_document(document, tl=tl)
    finally:
        if tl is not None:
            try:
                tl.close()
            except Exception:
                pass

    print("Nombre de mots :", cmpt)
    print("Nombre de corrections :", nb_corrections)

    chemin_corrige = chemin.with_name(chemin.stem + "_corrige.docx")
    try:
        document.save(str(chemin_corrige))
    except PermissionError:
        print(f"Impossible d'enregistrer {chemin_corrige} : fermez-le dans Word puis relancez.")
        sys.exit(1)
    print("Fichier corrigé enregistré :", chemin_corrige)
    if not ne_pas_ouvrir:
        ouvrir_fichier(chemin_corrige)


if __name__ == "__main__":
    main()