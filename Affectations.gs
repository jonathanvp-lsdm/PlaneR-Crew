/* =====================================================
   PLANE'R CREW
   AFFECTATIONS
===================================================== */

/* =====================================================
   ENREGISTRER UNE AFFECTATION
===================================================== */

function enregistrerAffectation(
    idBenevole,
    idPole
){

    const ss =
        SpreadsheetApp.getActiveSpreadsheet();

    const sheetAffectations =
        ss.getSheetByName("AFFECTATIONS");

    const sheetPoles =
        ss.getSheetByName("POLES");

    const sheetBenevoles =
        ss.getSheetByName("BENEVOLES");

    if(
        !sheetAffectations ||
        !sheetPoles ||
        !sheetBenevoles
    ){

        throw new Error(
            "Une des feuilles est introuvable."
        );

    }

    const affectations =
        sheetAffectations
            .getDataRange()
            .getValues();

    const poles =
        sheetPoles
            .getDataRange()
            .getValues();

    const benevoles =
        sheetBenevoles
            .getDataRange()
            .getValues();

    /* ============================================
       SUPPRESSION DES EN-TÊTES
    ============================================ */

    affectations.shift();
    poles.shift();
    benevoles.shift();

    /* ============================================
       VERIFICATION AFFECTATION EXISTANTE
    ============================================ */

    const affectationExistante =
        affectations.find(function(a){

            return (

                String(a[2]) === String(idBenevole)
                &&

                String(a[1]) === "2027"
                &&

                a[7] === "OUI"

            );

        });

    if(affectationExistante){

        const pole =
            poles.find(function(p){

                return Number(p[0]) === Number(affectationExistante[3]);

            });

        const benevole =
            benevoles.find(function(b){

                return String(b[0]) === String(idBenevole);

            });

        return{

            success:false,

            dejaAffecte:true,

            affectationId:
                affectationExistante[0],

            ancienPole:{

                id: pole[0],

                nom: pole[1],

                couleur: pole[5],

                icone: pole[6]

            },

            benevole:{

                id: benevole[0],

                nom: benevole[1],

                prenom: benevole[2]

            }

        };

    }

    /* ============================================
       ENREGISTREMENT
    ============================================ */

    const nouvelId =
        sheetAffectations.getLastRow();

    sheetAffectations.appendRow([

        nouvelId + 1,

        "2027",

        "'" + String(idBenevole).padStart(3,"0"),

        idPole,

        new Date(),

        getUtilisateurConnecte().nom,

        "",

        "OUI"

    ]);

    return{

        success:true,

        dejaAffecte:false

    };

}


/* =====================================================
   REAFFECTER UN BENEVOLE
===================================================== */

function reaffecterBenevole(
    affectationId,
    idBenevole,
    idPole
){

    const ss =
        SpreadsheetApp.getActiveSpreadsheet();

    const sheet =
        ss.getSheetByName("AFFECTATIONS");

    if(!sheet){
        throw new Error(
            "Feuille AFFECTATIONS introuvable."
        );
    }

    const data =
        sheet.getDataRange().getValues();

    /* Désactivation de l'ancienne affectation */

    for(let i = 1; i < data.length; i++){

        if(Number(data[i][0]) === Number(affectationId)){

            sheet
                .getRange(i + 1, 8)
                .setValue("NON");

            break;

        }

    }

    /* Création de la nouvelle */

    const nouvelId =
        sheet.getLastRow() + 1;

    sheet.appendRow([

        nouvelId,

        "2027",

        "'" + String(idBenevole).padStart(3,"0"),

        idPole,

        new Date(),

        getUtilisateurConnecte().nom,

        "Réaffectation automatique",

        "OUI"

    ]);

    validerBenevoleComplet(
    idBenevole
    );

    return{

        success:true,

        code:"REAFFECTATION_OK",

        message:"Réaffectation effectuée."

    };

}

/* =====================================================
   RECUPERER LES AFFECTATIONS ACTIVES
===================================================== */

function getAffectationsActives(){

    const ss =
        SpreadsheetApp.getActiveSpreadsheet();

    const sheetAffectations =
        ss.getSheetByName("AFFECTATIONS");

    const sheetBenevoles =
        ss.getSheetByName("BENEVOLES");

    const sheetPoles =
        ss.getSheetByName("POLES");

    const affectations =
        sheetAffectations
            .getDataRange()
            .getValues();

    const benevoles =
        sheetBenevoles
            .getDataRange()
            .getValues();

    const poles =
        sheetPoles
            .getDataRange()
            .getValues();

    affectations.shift();
    benevoles.shift();
    poles.shift();

    return affectations

        .filter(function(a){

            return a[7] === "OUI";

        })

        .map(function(a){

            const benevole =
                benevoles.find(
                    b => String(b[0]) === String(a[2])
                );

            const pole =
                poles.find(
                    p => Number(p[0]) === Number(a[3])
                );

            return{

                id:a[0],

                idBenevole:a[2],

                nom:
                    benevole
                        ? benevole[2] + " " + benevole[1]
                        : "Inconnu",

                pole:
                    pole
                        ? pole[6] + " " + pole[1]
                        : "Inconnu",

                idPole:a[3],

                affectePar:a[5],

                date: Utilities.formatDate(
                    new Date(a[4]),
                    Session.getScriptTimeZone(),
                    "dd/MM/yyyy HH:mm"
                )

            };

        });

}