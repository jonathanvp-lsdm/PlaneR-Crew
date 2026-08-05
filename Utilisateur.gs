/* =====================================================
   UTILISATEUR CONNECTE
===================================================== */

function getUtilisateurConnecte(){

    const email =
        Session.getActiveUser().getEmail();

    if(!email){

        throw new Error(
            "Impossible d'identifier l'utilisateur connecté."
        );

    }

    const sheet =
        SpreadsheetApp
            .getActive()
            .getSheetByName("UTILISATEURS");

    if(!sheet){

        throw new Error(
            "Feuille UTILISATEURS introuvable."
        );

    }

    const data =
        sheet
            .getDataRange()
            .getValues();

    data.shift();

    const utilisateur =
        data.find(function(ligne){

            return (

                String(ligne[0]).toLowerCase() ===
                email.toLowerCase()

                &&

                String(ligne[2]).toUpperCase() ===
                "OUI"

            );

        });

    if(!utilisateur){

        throw new Error(

            "Utilisateur non autorisé : " + email

        );

    }

    return {

    email: utilisateur[0],

    nom: utilisateur[1],

    actif: utilisateur[2]

};

}