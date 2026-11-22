import subprocess
import json
import sys
import bisect
import re
from docx import Document
from pathlib import Path  # bibliothèque pour retourner au chemin local du PC
import language_tool_python
import unicodedata
from wordfreq import zipf_frequency

try:
    with open("fautes_orthographes/fautes_generees.json", "r", encoding="utf-8") as f:
        FAUTES_GENEREES = json.load(f)
except FileNotFoundError:
    FAUTES_GENEREES = {}

"""
Cette fonction permet de corriger un texte avec l'outil LanguageTool,
sans toucher aux mots de la liste ignores (noms propres, termes techniques)
"""
def corriger_texte(texte, outil, mots_ignores):
    """Corrige un texte et renvoie (texte_corrigé, liste_des_corrections)."""
    matches = outil.check(texte)
    corrections = []

    # On part de la FIN pour que les positions des erreurs précédentes restent justes
    for match in reversed(matches):
        debut = match.offset                  # début de l'erreur EN COURS
        fin = debut + match.error_length      # fin de l'erreur EN COURS
        faute = texte[debut:fin]              # on découpe le TEXTE, pas match

        if faute in mots_ignores:             # nom propre ou terme technique → on ignore
            continue
        if not match.replacements:            # pas de suggestion → on laisse tel quel
            continue

        correction = match.replacements[0]
        texte = texte[:debut] + correction + texte[fin:]   # avant + correction + après
        corrections.append((faute, correction, match.message))

    corrections.reverse()   # remettre les corrections dans l'ordre de lecture
    return texte, corrections



def distance(a, b):
    """Nombre de lettres à ajouter, enlever ou changer pour passer de a à b (distance de Levenshtein)."""
    precedente = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        actuelle = [i]
        for j, cb in enumerate(b, 1):
            actuelle.append(min(precedente[j] + 1,                 # enlever
                                actuelle[j - 1] + 1,               # ajouter
                                precedente[j - 1] + (ca != cb)))   # changer
        precedente = actuelle
    return precedente[-1]


def phonetique(mot):
    """Transforme un mot en une 'clé de son' : simpatik et sympathique donnent la même clé."""
    mot = unicodedata.normalize("NFD", mot.lower())
    mot = "".join(c for c in mot if unicodedata.category(c) != "Mn")   # enlève les accents
    mot = re.sub(r"[^a-z]", "", mot)
    for avant, apres in [("eau", "o"), ("au", "o"), ("ph", "f"), ("th", "t"), ("qu", "k"),
                         ("c(?=[eiy])", "s"), ("g(?=[eiy])", "j"), ("c", "k"), ("y", "i"),
                         ("ai", "e"), ("ei", "e"), ("e?[ae]n(?=[^aeiou]|$)", "an"), ("z", "s")]:
        mot = re.sub(avant, apres, mot)
    mot = re.sub(r"(.)\1", r"\1", mot)        # lettres doublées → une seule
    mot = re.sub(r"[estxd]+$", "", mot)       # lettres muettes à la fin
    return mot


def choisir_suggestion(faute, suggestions):
    """Choisit la meilleure suggestion de LanguageTool, ou None si aucune n'est assez proche."""
    meilleure, meilleur_score = None, None
    for suggestion in suggestions[:5]:                         # les 5 premières suffisent
        # on compare sans apostrophes ni tirets : aujoudui est proche de aujourd'hui
        d = distance(re.sub(r"['’-]", "", faute.lower()), re.sub(r"['’-]", "", suggestion.lower()))
        meme_son = phonetique(faute) == phonetique(suggestion) and phonetique(faute) != ""
        if d > max(2, len(faute) // 4) and not meme_son:     # trop éloignée → on ne la prend pas
            continue
        score = zipf_frequency(suggestion, "fr") - d + (3 if meme_son else 0)
        if meilleur_score is None or score > meilleur_score:
            meilleure, meilleur_score = suggestion, score
    return meilleure


def motif_mot_entier(faute):
    return re.compile(r"(?<!\w)" + re.escape(faute) + r"(?!\w)")


def avec_majuscules(data):
    complet = dict(data)
    for faute, correction in data.items():
        if faute[:1].islower():
            complet.setdefault(faute[0].upper() + faute[1:], correction[0].upper() + correction[1:])
    return complet


def appliquer_correction(paragraphe, faute, correction):
    """Remplace `faute` par `correction` dans un paragraphe Word en gardant la mise en forme.
    Renvoie True si la correction a été appliquée, False sinon."""
    # Sortie anticipée : la faute n'est pas dans ce paragraphe
    if faute not in paragraphe.text:
        return False
    motif = motif_mot_entier(faute)
    if not motif.search(paragraphe.text):
        return False

    # Étape 1 : correction run par run (garde le style de chaque morceau)
    corrige = False
    for run in paragraphe.runs:  # chaque morceau du paragraphe
        if motif.search(run.text):
            run.text = motif.sub(lambda m: correction, run.text)
            corrige = True

    # Étape 2 : la faute est coupée entre plusieurs runs
    if not corrige:
        nouveau_texte = motif.sub(lambda m: correction, paragraphe.text)
        runs_texte = [run for run in paragraphe.runs if run.text]  # runs avec du texte (pas les images)
        runs_texte[0].text = nouveau_texte                         # tout le texte corrigé dans le 1er run
        for run in runs_texte[1:]:
            run.text = ""                                          # on vide les autres

    return True


def main():
    file_name = input("Quel est le nom de votre fichier à corriger (sans extension) : \n")

    home = Path.home()           # root path
    desktop = home / "Desktop"   # bureau
    resultats = list(desktop.rglob(file_name + ".docx"))  # liste de tous les fichiers trouvés portant ce nom
    if len(resultats) == 0:
        print("Impossible de trouver le fichier correspondant sur votre machine \n")
        print("Essayez une autre syntaxe")
        sys.exit()
    if len(resultats) > 1:
        print("Faites attention, il y a un autre fichier qui porte le même nom")

    path_file = resultats[0]
    document = Document(path_file)  # sert à parcourir les paragraphes

    # ouverture des JSON
    with open("fautes_orthographes/errors.json", "r", encoding='utf-8') as f1:
        data = json.load(f1)            # fautes et corrections personnelles

    with open("fautes_orthographes/ignorer.json", "r", encoding='utf-8') as f2:
        ignores = set(json.load(f2))    # mots à ne pas corriger

    tl = language_tool_python.LanguageTool('fr')  # créé UNE SEULE FOIS, avant la boucle

    cmpt, nb_corrections = corriger_document(document, data, ignores, tl)  # toute la correction

    tl.close()  # on ferme l'outil

    print(cmpt)
    print("Nombre de corrections :", nb_corrections)

    chemin_corrige = path_file.with_stem(path_file.stem + "_corrige")
    document.save(chemin_corrige)
    subprocess.run(["cmd", "/c", "start", "winword", str(chemin_corrige)])  # ouvre le fichier avec Word


def corriger_document(document, data, ignores, tl):
    cmpt = 0  # compteur de mots
    nb_corrections = 0  # compteur de corrections

    data = {**FAUTES_GENEREES, **avec_majuscules(data)}
    paragraphes = [p for p in document.paragraphs if p.text.strip()]  # paragraphe vide → on passe au suivant

    # Passe 1 : on cherche dans le fichier avec les fautes et les corrections
    for paragraphe in paragraphes:
        texte = paragraphe.text
        cmpt += len(texte.split())
        for faute, correction in data.items():
            if faute in ignores or faute not in texte:
                continue
            if appliquer_correction(paragraphe, faute, correction):
                nb_corrections += 1
                texte = paragraphe.text

    if tl is None:
        return cmpt, nb_corrections

    # Passe 2 : LanguageTool, 2 tours : certaines fautes n'apparaissent qu'une fois les premières corrigées
    for tour in range(2):
        textes = [p.text for p in paragraphes]
        texte_complet = "\n".join(textes)
        matches = tl.check(texte_complet)

        debuts = []
        position = 0
        for texte in textes:
            debuts.append(position)
            position += len(texte) + 1

        corrige_ce_tour = 0
        for match in matches:
            faute = texte_complet[match.offset : match.offset + match.error_length]
            if faute in ignores or not match.replacements:
                continue

            correction = choisir_suggestion(faute, match.replacements)
            if correction is None:      # aucune suggestion assez proche → mieux vaut laisser la faute
                print(f"{faute} : laissé tel quel ({match.replacements[:3]})")
                continue

            i = bisect.bisect_right(debuts, match.offset) - 1
            if appliquer_correction(paragraphes[i], faute, correction):
                nb_corrections += 1
                corrige_ce_tour += 1
                print(f"{faute} -> {correction} ({match.message})")

        if corrige_ce_tour == 0:        # rien de nouveau → pas besoin d'un 2e tour
            break

    return cmpt, nb_corrections


if __name__ == "__main__":
    main()