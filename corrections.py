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
except ImportError:                       # wordfreq absent -> on fait sans
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
                outil = language_tool_python.LanguageTool("fr")
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
# Correction complète
# ---------------------------------------------------------------------------

def corriger_document(document, data=None, ignores=None, tl=None, infos=None):
    """Corrige le document (modifié sur place). Renvoie (nombre_de_mots, nombre_de_corrections).
    tl = outil LanguageTool, ou None pour corriger seulement avec les dictionnaires.
    infos (dict, optionnel) reçoit "mode" : "local", "public" ou "dictionnaires"."""
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