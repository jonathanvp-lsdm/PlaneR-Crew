// =====================================================
// PLANE'R CREW
// Inscriptions.gs
// Historique bénévole + IDs forcés en texte
// =====================================================


function creerCandidatureBenevole(donnees){

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetBenevoles = ss.getSheetByName("BENEVOLES");
  const sheetInscriptions = ss.getSheetByName("INSCRIPTIONS");

  if(!sheetBenevoles || !sheetInscriptions){
    throw new Error("Onglet BENEVOLES ou INSCRIPTIONS introuvable");
  }

  const data = sheetBenevoles.getDataRange().getValues();
  const headers = data[0];

  const colEmail = headers.indexOf("EMAIL");
  const colId = headers.indexOf("ID");
  const colNb = headers.indexOf("NB_CANDIDATURES");
  const colEdition = headers.indexOf("DERNIERE_EDITION");
  const colStatut = headers.indexOf("DERNIER_STATUT");
  const colDate = headers.indexOf("DATE_DERNIERE_CANDIDATURE");

  let idBenevole = null;
  let ligneBenevole = null;


  for(let i=1;i<data.length;i++){

    if(String(data[i][colEmail]).toLowerCase() === String(donnees.EMAIL).toLowerCase()){

      idBenevole = String(data[i][colId]).padStart(3,"0");
      ligneBenevole = i+1;
      break;

    }

  }


  if(!idBenevole){

    idBenevole = genererNouvelIDBenevole();

    const ligne = sheetBenevoles.getLastRow()+1;

    sheetBenevoles.getRange(ligne,1,1,15)
      .setNumberFormat("@");

    sheetBenevoles.getRange(ligne,1,1,15)
      .setValues([[

        idBenevole,
        donnees.NOM || "",
        donnees.PRENOM || "",
        donnees.EMAIL || "",
        donnees.TELEPHONE || "",
        "",
        "",
        donnees.TAILLE_TSHIRT || "",
        "NON",
        new Date(),
        new Date(),
        1,
        donnees.ID_EDITION || "",
        "A_TRAITER",
        new Date()

      ]]);

  }


  else {

    sheetBenevoles.getRange(ligneBenevole,colNb+1)
      .setValue(Number(data[ligneBenevole-1][colNb] || 0)+1);

    sheetBenevoles.getRange(ligneBenevole,colEdition+1)
      .setValue(donnees.ID_EDITION || "");

    sheetBenevoles.getRange(ligneBenevolenevole,colStatut+1)
      .setValue("A_TRAITER");

    sheetBenevoles.getRange(ligneBenevole,colDate+1)
      .setValue(new Date());

  }


  const idInscription = genererNouvelIDInscription();

  const ligneInscription = sheetInscriptions.getLastRow()+1;

  sheetInscriptions.getRange(ligneInscription,1,1,9)
    .setNumberFormat("@");

  sheetInscriptions.getRange(ligneInscription,1,1,9)
    .setValues([[

      idInscription,
      idBenevole,
      donnees.ID_EDITION || "",
      "A_TRAITER",
      donnees.CHOIX_1 || "",
      donnees.CHOIX_2 || "",
      donnees.CHOIX_3 || "",
      donnees.CHOIX_4 || "",
      donnees.CHOIX_5 || ""

    ]]);


  clearBenevolesCache();

  return {
    success:true,
    idBenevole:idBenevole
  };

}



function genererNouvelIDBenevole(){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BENEVOLES");

  const ids = sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,1),1)
    .getValues()
    .flat()
    .map(Number)
    .filter(n=>!isNaN(n));

  return String((ids.length ? Math.max(...ids) : 0)+1).padStart(3,"0");

}



function genererNouvelIDInscription(){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INSCRIPTIONS");

  const ids = sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,1),1)
    .getValues()
    .flat()
    .map(Number)
    .filter(n=>!isNaN(n));

  return String((ids.length ? Math.max(...ids) : 0)+1).padStart(3,"0");

}



/* =====================================================
   TEST CREATION CANDIDATURE
===================================================== */

function testCreerCandidature(){

  const resultat = creerCandidatureBenevole({

    NOM:"TESTEUR",
    PRENOM:"CRM",
    EMAIL:"testeur6.crm@planerfest.fr",
    TELEPHONE:"0600000000",
    TAILLE_TSHIRT:"L",
    ID_EDITION:"2028",
    CHOIX_1:"BAR",
    CHOIX_2:"VIP"

  });


  console.log(resultat);

}
