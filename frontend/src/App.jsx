// 1er code react !
// composant = fonction

import {useState} from "react"; // etat permet de conccdevoir les données
// si une variable pas de MAJ merci aux états

export default function App(){
   const [fichier,setFichier] = useState(null);// permet de mémo une donnée

    function modifieretat(event){
    const fichier_select =  event.target.files[0] // on recup le fichier
        setFichier(fichier_select)
        // tests
        console.log(fichier_select)
    }

    return  (   <>
            <h1>Correcteur de documents Word (en français)</h1>
            {/* accepte seulement les fichiers Word */}
            <input  onChange={modifieretat} type="file" accept=".docx" />
        </>
    )

}

