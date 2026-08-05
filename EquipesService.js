/* =====================================================
   PLANE'R CREW
   MODULE EQUIPES
===================================================== */

function getPageEquipes(){

  return HtmlService
    .createHtmlOutputFromFile("Equipes")
    .getContent();

}

/* =====================================================
   RECUPERATION DES POLES
===================================================== */

function getEquipes(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("POLES");

  if(!sheet){

    throw new Error("Feuille POLES introuvable.");

  }

  const data = sheet.getDataRange().getValues();

  const headers = data.shift();

 const affectations =
    ss.getSheetByName("AFFECTATIONS")
      .getDataRange()
      .getValues();

return data

    .filter(ligne => ligne[2] === "OUI")

    .map(ligne => {

        const idPole = ligne[0];

        const nbAffectes =
            affectations.filter(a =>

                Number(a[3]) === Number(idPole)
                &&
                a[7] === "OUI"

            ).length;

        return {

            id: ligne[0],
            nom: ligne[1],
            actif: ligne[2],
            ordre: ligne[3],
            nbBenevoles: ligne[4],
            couleur: ligne[5],
            icone: ligne[6],
            nbAffectes: nbAffectes

        };

    })

    .sort((a,b)=>a.ordre-b.ordre);

    }
  