import subprocess
import json
import sys
from docx import Document
from pathlib import Path  # bibliothèque pour retourner au chemin local du PC
import language_tool_python

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


def appliquer_correction(paragraphe, faute, correction):
    """Remplace `faute` par `correction` dans un paragraphe Word en gardant la mise en forme.
    Renvoie True si la correction a été appliquée, False sinon."""

    # Sortie anticipée : la faute n'est pas dans ce paragraphe
    if faute not in paragraphe.text:
        return False

    # Étape 1 : correction run par run (garde le style de chaque morceau)
    corrige = False
    for run in paragraphe.runs:  # chaque morceau du paragraphe
        if faute in run.text:
            run.text = run.text.replace(faute, correction)
            corrige = True

    # Étape 2 : la faute est coupée entre plusieurs runs
    if not corrige:
        nouveau_texte = paragraphe.text.replace(faute, correction)
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

    cmpt = 0              # compteur de mots
    nb_corrections = 0    # compteur de corrections

    for paragraphe in document.paragraphs:
        if not paragraphe.text.strip():   # paragraphe vide → on passe au suivant
            continue
        cmpt += len(paragraphe.text.split())

        # Passe 1 : on cherche dans le fichier avec les fautes et les corrections
        for faute, correction in data.items():
            if appliquer_correction(paragraphe, faute, correction):
                nb_corrections += 1

        # Passe 2 : LanguageTool sur le texte déjà corrigé par la passe 1
        _, corrections_lt = corriger_texte(paragraphe.text, tl, ignores)
        for faute, correction, message in corrections_lt:
            if appliquer_correction(paragraphe, faute, correction):
                nb_corrections += 1
                print(f"{faute} -> {correction} ({message})")

    tl.close()   # on ferme l'outil

    print(cmpt)
    print("Nombre de corrections :", nb_corrections)

    chemin_corrige = path_file.with_stem(path_file.stem + "_corrige")
    document.save(chemin_corrige)
    subprocess.run(["cmd", "/c", "start", "winword", str(chemin_corrige)])  # ouvre le fichier avec Word


if __name__ == "__main__":
    main()