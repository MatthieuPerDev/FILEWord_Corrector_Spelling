// 1er code react !
// composant = fonction

import {useState,useRef,useEffect} from "react"; // etat permet de conccdevoir les données
// si une variable pas de MAJ merci aux états
const API_URL = import.meta.env.VITE_API_URL || "http://localhost:8000"
import {renderAsync} from "docx-preview"
export default function App() {
    const [fichier, setFichier] = useState(null);// permet de mémo une donnée
    const [fichierCorrige, setFichierCorrige] = useState(null)
    const [enCours, setenCours] = useState(false)
    const [erreur, seterreur] = useState(null)
    const [resultat, setResultat] = useState(null)   // nombre de corrections + mode utilisé

    const apercuAvant = useRef(null)    // la <div> où on dessine l'original
    const apercuApres = useRef(null)    // la <div> où on dessine le corrigé

    useEffect(() => {
        if (fichier && apercuAvant.current) {
            renderAsync(fichier, apercuAvant.current, null, { ignoreWidth: true })
        }
    }, [fichier])

    // quand le fichier corrigé arrive, on dessine l'aperçu "après"
    useEffect(() => {
        if (fichierCorrige && apercuApres.current) {
            renderAsync(fichierCorrige, apercuApres.current, null, { ignoreWidth: true })
        }
    }, [fichierCorrige])

    function modifieretat(event) {
        const fichier_select = event.target.files[0] // on recup le fichier
        setFichier(fichier_select)
        setFichierCorrige(null)    // ← on efface l'ancien résultat
        setResultat(null)
        /*
            console.log(fichier_select)
    }
*/}
    async function envoyerFichier() {

        const formData = new FormData()
        formData.append("file", fichier)
        seterreur(null)
        setResultat(null)
        setenCours(true)

        try {
            const response = await fetch(`${API_URL}/corriger`, {
                method: "POST",
                body: formData,
            })
            if (!response.ok) {
                const data = await response.json()     // l'API renvoie {"detail": "..."}
                throw new Error(data.detail)
            }
            setResultat({
                statut: response.headers.get("X-Correction"),
                nb: response.headers.get("X-Nb-Corrections"),
            })
            const blob = await response.blob()
            setFichierCorrige(blob)

            const url = URL.createObjectURL(blob)
            const lien = document.createElement("a")
            lien.href = url
            lien.download = fichier.name.replace(".docx", "_corrige.docx")
            lien.click()


            URL.revokeObjectURL(url)

        } catch (e) {
            seterreur(e.message)


        } finally {
            setenCours(false)

        }


    }

    return (


        <div className="min-h-screen bg-emerald-100 flex flex-col items-center justify-center gap-6 p-8">
            <div className="bg-white rounded-xl shadow-lg p-8 max-w-md w-full flex flex-col gap-4">

                <h1 className="text-3xl font-bold text-blue-600 text-center">Correcteur de documents Word (en
                    français)</h1>

                {/* accepte seulement les fichiers Word */}
                <input
                    className="file:bg-blue-500 file:text-white file:px-4 file:py-2 file:rounded-lg file:border-0 hover:file:bg-yellow-500 file:transition"
                    onChange={modifieretat} type="file" accept=".docx"/>

                {fichier != null ? (
                    <div className="text-center">
                        <p>Nom du fichier : {fichier.name}</p>
                        <p>Taille du fichier : {(fichier.size / 1024.).toFixed(1)} Ko</p>
                    </div>
                ) : (
                    <p className="text-center text-gray-500">Sélectionnez un fichier Word (.docx) à corriger</p>
                )}

                <button
                    className="bg-cyan-600 text-white py-2 rounded-lg hover:bg-yellow-500 transition disabled:bg-gray-400 disabled:cursor-not-allowed"
                    disabled={fichier == null || enCours} onClick={envoyerFichier}>
                    Corriger
                </button>

                {enCours && <p className="text-center text-blue-600">Correction en cours ...</p>}
                {erreur && <p className="text-center text-red-600">Erreur : {erreur}</p>}
                {resultat && (
                    <p className={`text-center ${resultat.statut === "partielle" ? "text-orange-600" : "text-green-700"}`}>
                        {resultat.nb ?? "?"} correction(s) — {{
                        ia: "par l'IA",
                        complete: "dictionnaires + LanguageTool",
                        partielle: "dictionnaires seulement (LanguageTool indisponible sur le serveur)",
                    }[resultat.statut] ?? resultat.statut}
                    </p>
                )}
            </div>

            {fichier && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 w-full max-w-7xl">

                    <div className="bg-white rounded-xl shadow-lg p-4">
                        <h2 className="text-xl font-bold text-red-500 mb-2">Avant</h2>
                        <div ref={apercuAvant} className="overflow-auto max-h-[250px] [zoom:0.8]"></div>
                    </div>

                    <div className="bg-white rounded-xl shadow-lg p-4">
                        <h2 className="text-xl font-bold text-green-600 mb-2">Après</h2>
                        {fichierCorrige
                            ? <div ref={apercuApres} className="overflow-auto max-h-[250px] [zoom:0.8] "></div>
                            : <p className="text-gray-500">Cliquez sur Corriger</p>}
                    </div>

                </div>
            )}



        </div>
    )
}