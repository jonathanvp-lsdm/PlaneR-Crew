/* =====================================================
   PLANE'R CREW
   VALIDATION BENEVOLES
===================================================== */


/* =====================================================
   VALIDATION COMPLETE BENEVOLE
===================================================== */

function validerBenevoleComplet(idBenevole){

  console.log("ID RECU :", idBenevole);

  if(!idBenevole){
    throw new Error("ID BENEVOLE MANQUANT");
  }

  idBenevole = String(idBenevole).padStart(3,"0");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetIns = ss.getSheetByName("INSCRIPTIONS");
  const sheetBen = ss.getSheetByName("BENEVOLES");

  if(!sheetIns || !sheetBen){
    throw new Error("Onglet introuvable");
  }


  // ============================
  // VALIDATION INSCRIPTION
  // ============================

  const dataIns = sheetIns.getDataRange().getValues();
  const headersIns = dataIns[0].map(String);

  const colIdBen = headersIns.indexOf("ID_BENEVOLE");
  const colStatut = headersIns.indexOf("STATUT");

  for(let i=1;i<dataIns.length;i++){

    if(String(dataIns[i][colIdBen]).padStart(3,"0")===idBenevole){

      sheetIns
        .getRange(i+1,colStatut+1)
        .setValue("VALIDE");

      break;

    }

  }


  // ============================
  // HISTORIQUE BENEVOLE
  // ============================

  const dataBen = sheetBen.getDataRange().getValues();
  const headersBen = dataBen[0].map(String);

  const colId = headersBen.indexOf("ID");
  const colActif = headersBen.indexOf("EST_ACTIF");
  const colDernierStatut = headersBen.indexOf("DERNIER_STATUT");
  const colDate = headersBen.indexOf("DATE_DERNIERE_CANDIDATURE");

  for(let i=1;i<dataBen.length;i++){

    if(String(dataBen[i][colId]).padStart(3,"0")===idBenevole){

      if(colActif>-1)
        sheetBen.getRange(i+1,colActif+1).setValue("OUI");

      if(colDernierStatut>-1)
        sheetBen.getRange(i+1,colDernierStatut+1).setValue("VALIDE");

      if(colDate>-1)
        sheetBen.getRange(i+1,colDate+1).setValue(new Date());

      break;

    }

  }

  clearBenevolesCache();

  return true;

}



/* =====================================================
   CHANGEMENT DE STATUT
===================================================== */

function updateStatutBenevole(idBenevole,statut){

  idBenevole = String(idBenevole).padStart(3,"0");

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("INSCRIPTIONS");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);

  const colId = headers.indexOf("ID_BENEVOLE");
  const colStatut = headers.indexOf("STATUT");

  for(let i=1;i<data.length;i++){

    if(String(data[i][colId]).padStart(3,"0")===idBenevole){

      sheet
        .getRange(i+1,colStatut+1)
        .setValue(statut);

      break;

    }

  }

  clearBenevolesCache();

  return true;

}

/* =====================================================
   VALIDER ET AFFECTER UN BENEVOLE
===================================================== */

function validerEtAffecterBenevole(
  idBenevole,
  idPole,
  affectePar
){

  enregistrerAffectation(
    idBenevole,
    idPole,
    affectePar
  );

  validerBenevoleComplet(
    idBenevole
  );

  return true;

}

/* =====================================================
   TEST
===================================================== */

function testValidation004(){

  validerBenevoleComplet("004");

}