// 1er code react !
// composant = fonction

import {useState} from "react"; // etat permet de conccdevoir les données
// si une variable pas de MAJ merci aux états

export default function App(){
   const [fichier,setFichier] = useState(null);// permet de mémo une donnée
    const [enCours,setenCours] = useState(false)
    const [erreur,seterreur] = useState(null)
    function modifieretat(event){
    const fichier_select =  event.target.files[0] // on recup le fichier
        setFichier(fichier_select)
        // tests
        console.log(fichier_select)
    }

    async function envoyerFichier(){

        const formData = new FormData()
        formData.append("file",fichier)
        seterreur(null)
        setenCours(true)

try {
    const response = await fetch("http://localhost:8000/corriger", {
        method: "POST",
        body: formData,
    })
 if(!response.ok) {
        throw new Error("Le fichier est corrompu !")
    }
    const blob = await response.blob()

    const url = URL.createObjectURL(blob)
    const lien = document.createElement("a")
    lien.href = url
    lien.download = fichier.name.replace(".docx", "_corrige.docx")
    lien.click()


    URL.revokeObjectURL(url)

}
    catch (e) {
        seterreur(e.message)


    }
        finally {
    setenCours(false)

}


    }

    return  (   <>
            <h1>Correcteur de documents Word (en français)</h1>
            {/* accepte seulement les fichiers Word */}
            <input  onChange={modifieretat} type="file" accept=".docx" />
            {fichier != null ? (
            <div>
                <p>Nom du fichier :  {fichier.name}</p>
                <p> Taille du fichier : {(fichier.size/1024.).toFixed(1)} Ko</p>
            </div>


                )
    : ( <p>Aucun fichier choisi ! </p>
    )}
            <button disabled={fichier == null ||enCours} onClick={envoyerFichier}> Corriger</button>
            {enCours &&
            <p>Correction en cours ... </p>}
            {erreur &&
            <p> Erreur : {erreur} </p>}
        </>

    )

}

