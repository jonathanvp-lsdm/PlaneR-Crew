/* =====================================================
   PLANE'R CREW
   AFFECTATIONS
===================================================== */

/* =====================================================
   ENREGISTRER UNE AFFECTATION
===================================================== */

function enregistrerAffectation(
    idBenevole,
    idPole,
    affectePar
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

        affectePar,

        "",

        "OUI"

    ]);

    return{

        success:true,

        dejaAffecte:false

    };

}

/* =====================================================
   TEST
===================================================== */

function testAffectation(){

    Logger.log(

        enregistrerAffectation(

            "001",

            1,

            "Alex"

        )

    );

}