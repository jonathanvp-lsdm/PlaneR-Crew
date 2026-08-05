/* =====================================================
   PLANE'R CREW
   DETAIL D'UNE EQUIPE
===================================================== */

function getEquipe(idPole){

    const ss =
        SpreadsheetApp.getActiveSpreadsheet();

    const sheetPoles =
        ss.getSheetByName("POLES");

    const sheetAffectations =
        ss.getSheetByName("AFFECTATIONS");

    const sheetBenevoles =
        ss.getSheetByName("BENEVOLES");

    if(
        !sheetPoles ||
        !sheetAffectations ||
        !sheetBenevoles
    ){
        throw new Error("Une des feuilles est introuvable.");
    }

    const poles =
        sheetPoles
            .getDataRange()
            .getValues();

    poles.shift();

    const affectations =
        sheetAffectations
            .getDataRange()
            .getValues();

    affectations.shift();

    const benevoles =
        sheetBenevoles
            .getDataRange()
            .getValues();

    benevoles.shift();

    /* ============================================
       RECHERCHE DU POLE
    ============================================ */

    const pole =
        poles.find(function(ligne){

            return Number(ligne[0]) === Number(idPole);

        });

    if(!pole){

        throw new Error("Pôle introuvable.");

    }

    /* ============================================
       AFFECTATIONS DU POLE
    ============================================ */

    const affectationsPole =
        affectations.filter(function(a){

            return (
                Number(a[3]) === Number(idPole)
                &&
                a[7] === "OUI"
            );

        });

    /* ============================================
       BENEVOLES
    ============================================ */

    const listeBenevoles =
        affectationsPole.map(function(a){

            const benevole =
                benevoles.find(function(b){

                    return String(b[0]) === String(a[2]);

                });

            if(!benevole){
                return null;
            }

            return{

                id: benevole[0],

                nom: benevole[1],

                prenom: benevole[2],

                email: benevole[3],

                telephone: benevole[4],

                photoId: benevole[5],

                photoNom: benevole[6],

                tailleTshirt: benevole[7],

                actif: benevole[8]

            };

        }).filter(Boolean);

    /* ============================================
       RETOUR
    ============================================ */

    return{

        id: pole[0],

        nom: pole[1],

        actif: pole[2],

        ordre: pole[3],

        nbBenevoles: pole[4],

        couleur: pole[5],

        icone: pole[6],

        nbAffectes: listeBenevoles.length,

        benevoles: listeBenevoles

    };

}