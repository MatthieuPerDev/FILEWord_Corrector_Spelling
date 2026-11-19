import subprocess
import os
import json
import sys
from docx import Document
from pathlib import Path #bibliothéque pour retourner au chemin local du PC

file_name = input("Quelle est le nom de votre fichier à corriger(sans extension) : \n")

home = Path.home()       #root path
desktop = home / "Desktop"   # bureau
resultats = list(desktop.rglob(file_name + ".docx"))# liste de tous les fichiers trouvés portant ce no
if len(resultats) == 0 :
    print("Impossible de trouver le fichier correspondant sur votre machine \n")
    print("Essayer une autre syntaxe")
    sys.exit()
if len(resultats) >1 :
    print("Faites attention il y a un fichier qui porte le meme nom")

path_file = resultats[0]
document = Document(path_file) # sert à parcourir les parataphes
#ouverture du JSON

with open ('fautes_orthographes/errors.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

cmpt= 0 #nbr de mots
nb_corrections= 0 #nombre de correcrions

#Compte_rendu_entretien_Matthieu
for paragraphe in document.paragraphs:
    cmpt += len(paragraphe.text.split())

    for faute, correction in data.items():
        if faute not in paragraphe.text:   # faute absente de ce paragraphe → suivante
            continue

        # Étape 1 : correction run par run (garde la mise en forme)
        corrige = False
        for run in paragraphe.runs:
            if faute in run.text:
                run.text = run.text.replace(faute, correction)
                corrige = True

        # Étape 2 : la faute est coupée entre plusieurs runs
        if not corrige:
            nouveau_texte = paragraphe.text.replace(faute, correction)
            runs_texte = [run for run in paragraphe.runs if run.text]  # runs avec du texte (pas les images)
            runs_texte[0].text = nouveau_texte                         # texte corrigé dans le 1er run
            for run in runs_texte[1:]:
                run.text = ""                                          # on vide les autres

        nb_corrections += 1

chemin_corrige = path_file.with_stem("Document_modifié")
document.save(chemin_corrige)
print("Fichier enregistré :", chemin_corrige)
subprocess.run(["cmd", "/c", "start", "winword", str(chemin_corrige)])
