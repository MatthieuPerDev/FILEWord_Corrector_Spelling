"""
Génère automatiquement une liste de fautes d'orthographe courantes en français.

Principe :
1. On prend les mots français les plus fréquents (wordfreq).
2. On fabrique des fautes typiques : accent oublié, mauvais accent,
   consonne double oubliée, consonne doublée à tort.
3. On ne garde une faute QUE si elle n'existe dans aucun dictionnaire
   (français ni anglais) et si elle ne peut venir que d'un seul mot.
   → une telle faute est TOUJOURS fausse, quel que soit le contexte.
"""
import json
import sys
from collections import defaultdict
from wordfreq import top_n_list, zipf_frequency
from spellchecker import SpellChecker

NB_MOTS_SOURCE = int(sys.argv[1]) if len(sys.argv) > 1 else 20000
NB_FAUTES_MAX = int(sys.argv[2]) if len(sys.argv) > 2 else 2000

fr = SpellChecker(language="fr")
en = SpellChecker(language="en")

# pas de î ni de û : depuis la réforme de 1990, "connait", "maitre", "surement" sont corrects
ACCENTS = {"é": "e", "è": "e", "ê": "e", "à": "a", "â": "a",
           "ô": "o", "ù": "u", "ç": "c"}
ECART_MIN = 1.0   # la faute doit être au moins 10 fois plus rare que le bon mot
# pas de é -> è : depuis 1990, "évènement" et "règlementation" sont corrects
ECHANGES = {"é": ["ê"], "è": ["é"], "ê": ["é", "è"]}
DOUBLES = ["mm", "nn", "ll", "tt", "ss", "rr", "pp", "ff", "cc"]
VOYELLES = set("aeiouyéèêàâîôû")


# formes rectifiées de 1990 (sans î ni û) : elles sont correctes aussi
RECTIFIEES = {m.replace("î", "i").replace("û", "u") for m in fr.word_frequency.dictionary}


def est_un_vrai_mot(mot):
    return mot in fr or mot in en or mot in RECTIFIEES


def variantes(mot):
    """Toutes les fautes typiques qu'on peut faire sur ce mot."""
    resultats = set()
    for i, lettre in enumerate(mot):
        # accent oublié : déjà -> deja (une lettre à la fois)
        if lettre in ACCENTS:
            resultats.add(mot[:i] + ACCENTS[lettre] + mot[i + 1:])
        # mauvais accent : succès -> succés
        for autre in ECHANGES.get(lettre, []):
            resultats.add(mot[:i] + autre + mot[i + 1:])
    # consonne double oubliée : développement -> dévelopement
    for double in DOUBLES:
        debut = mot.find(double)
        while debut != -1:
            resultats.add(mot[:debut] + mot[debut + 1:])
            debut = mot.find(double, debut + 1)
    # consonne doublée à tort entre deux voyelles : appeler <- apeller
    for i in range(1, len(mot) - 1):
        c = mot[i]
        if c in "mnltrpf" and mot[i - 1] in VOYELLES and mot[i + 1] in VOYELLES:
            resultats.add(mot[:i] + c + mot[i:])
    resultats.discard(mot)
    return resultats


# 1. les mots sources : fréquents, valides, assez longs, seulement des lettres
mots = [m for m in top_n_list("fr", NB_MOTS_SOURCE)
        if len(m) >= 5 and m.isalpha() and m in fr]

# 2. on fabrique les fautes des mots sources
fautes_sources = {}
for mot in mots:
    for faute in variantes(mot):
        if not est_un_vrai_mot(faute):
            fautes_sources.setdefault(faute, set()).add(mot)

# 3. ambiguïté vérifiée sur TOUT le dictionnaire : une faute qui peut venir
#    de deux mots différents (jete : jeté ou jette ?) est écartée
origines = defaultdict(set)
for mot in fr.word_frequency.dictionary:
    if mot.isalpha():
        for faute in variantes(mot):
            if faute in fautes_sources:
                origines[faute].add(mot)

candidats = [(faute, next(iter(src))) for faute, src in fautes_sources.items()
             if len(src) == 1 and len(origines[faute]) == 1]

# 3 bis. on écarte les "fautes" presque aussi fréquentes que le bon mot :
#        ce sont souvent de vrais mots absents du dictionnaire, ou des noms propres
#        (baise / baisse, gaulle / gaule, benoit / benoît)
candidats = [(f, c) for f, c in candidats
             if zipf_frequency(c, "fr") - zipf_frequency(f, "fr") >= ECART_MIN]

# 4. on classe : d'abord les mots les plus fréquents, et les fautes qu'on voit réellement
#    sur le web (wordfreq connaît aussi les fautes fréquentes)
candidats.sort(key=lambda fc: (zipf_frequency(fc[0], "fr"), zipf_frequency(fc[1], "fr")), reverse=True)

resultat = dict(candidats[:NB_FAUTES_MAX])
json.dump(resultat, open("fautes_generees.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)

print("mots sources :", len(mots))
print("fautes candidates (non ambiguës) :", len(candidats))
print("fautes gardées :", len(resultat))
print("exemples :", list(resultat.items())[:40])