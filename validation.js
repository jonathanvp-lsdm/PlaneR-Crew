/* =====================================================
   PLANE'R CREW
   VALIDATION BENEVOLES
   VERSION CORRIGEE
===================================================== */


/* =====================================================
   RECUPERATION BENEVOLES + INSCRIPTIONS
===================================================== */

function getBenevoles(){

  const cache = CacheService.getScriptCache();

  const cached = cache.get("BENEVOLES_CACHE");


  if(cached){

  console.log("CACHE UTILISE");

  return JSON.parse(cached);

}

console.log("LECTURE GOOGLE SHEETS");

  const ss = SpreadsheetApp.getActiveSpreadsheet();


  const sheetBenevoles =
    ss.getSheetByName("BENEVOLES");


  const sheetInscriptions =
    ss.getSheetByName("INSCRIPTIONS");



  if(!sheetBenevoles || !sheetInscriptions){
    throw new Error(
      "Onglet BENEVOLES ou INSCRIPTIONS introuvable"
    );
  }



  const benevolesData =
    sheetBenevoles.getDataRange().getValues();


  const inscriptionsData =
    sheetInscriptions.getDataRange().getValues();



  const headersBen =
    benevolesData[0].map(function(h){
      return String(h).trim();
    });



  const headersIns =
    inscriptionsData[0].map(function(h){
      return String(h).trim();
    });



  let benevoles = {};



  for(let i=1;i<benevolesData.length;i++){

    let obj={};


    headersBen.forEach(function(header,index){

      obj[header] =
        benevolesData[i][index];

    });



    benevoles[
      String(obj.ID).padStart(3,"0")
    ] = obj;


  }





  let result=[];



  for(let i=1;i<inscriptionsData.length;i++){


    let inscription={};


    headersIns.forEach(function(header,index){

      inscription[header] =
        inscriptionsData[i][index];

    });



    let id =
      String(inscription.ID_BENEVOLE)
      .padStart(3,"0");



    result.push({

      ...benevoles[id],
      ...inscription,

      ID:
        benevoles[id].ID,

      ID_INSCRIPTION:
        inscription.ID

    });


  }



  cache.put(
    "BENEVOLES_CACHE",
    JSON.stringify(result),
    300
  );



  return result;


}







function clearBenevolesCache(){

  CacheService
    .getScriptCache()
    .remove(
      "BENEVOLES_CACHE"
    );

}








/* =====================================================
   VALIDATION COMPLETE BENEVOLE
===================================================== */


function validerBenevoleComplet(idBenevole){


  console.log(
    "ID RECU :",
    idBenevole
  );



  if(!idBenevole){

    throw new Error(
      "ID BENEVOLE MANQUANT"
    );

  }



  idBenevole =
    String(idBenevole)
    .padStart(3,"0");



  console.log(
    "ID NORMALISE :",
    idBenevole
  );



  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  const sheetIns =
    ss.getSheetByName("INSCRIPTIONS");


  const sheetBen =
    ss.getSheetByName("BENEVOLES");



  if(!sheetIns || !sheetBen){
    throw new Error(
      "Onglet introuvable"
    );
  }






/* =====================================================
   1 - VALIDATION INSCRIPTION
===================================================== */


  const dataIns =
    sheetIns.getDataRange().getValues();


  const headersIns =
    dataIns[0].map(function(h){
      return String(h).trim();
    });



  const colIdBen =
    headersIns.indexOf(
      "ID_BENEVOLE"
    );


  const colStatut =
    headersIns.indexOf(
      "STATUT"
    );



  for(let i=1;i<dataIns.length;i++){


    if(

      String(dataIns[i][colIdBen])
      .padStart(3,"0")

      ===

      idBenevole

    ){


      sheetIns
        .getRange(
          i+1,
          colStatut+1
        )
        .setValue(
          "VALIDE"
        );


      console.log(
        "INSCRIPTION VALIDEE LIGNE",
        i+1
      );


      break;

    }

  }







/* =====================================================
   2 - MISE A JOUR HISTORIQUE BENEVOLE
===================================================== */


  const dataBen =
    sheetBen.getDataRange().getValues();



  const headersBen =
    dataBen[0].map(function(h){
      return String(h).trim();
    });



  const colId =
    headersBen.indexOf(
      "ID"
    );


  const colActif =
    headersBen.indexOf(
      "EST_ACTIF"
    );


  const colDernierStatut =
    headersBen.indexOf(
      "DERNIER_STATUT"
    );


  const colDateDerniere =
    headersBen.indexOf(
      "DATE_DERNIERE_CANDIDATURE"
    );



  console.log(
    "COL ID :",
    colId
  );

  console.log(
    "COL ACTIF :",
    colActif
  );

  console.log(
    "COL DERNIER STATUT :",
    colDernierStatut
  );

  console.log(
    "COL DATE DERNIERE :",
    colDateDerniere
  );




  for(let i=1;i<dataBen.length;i++){



    let idSheet =
      String(dataBen[i][colId])
      .padStart(3,"0");



    console.log(
      "COMPARAISON :",
      idSheet,
      idBenevole
    );



    if(idSheet === idBenevole){



      if(colActif !== -1){

        sheetBen
          .getRange(
            i+1,
            colActif+1
          )
          .setValue(
            "OUI"
          );

      }




      if(colDernierStatut !== -1){

        sheetBen
          .getRange(
            i+1,
            colDernierStatut+1
          )
          .setValue(
            "VALIDE"
          );

      }




      if(colDateDerniere !== -1){

        sheetBen
          .getRange(
            i+1,
            colDateDerniere+1
          )
          .setValue(
            new Date()
          );

      }



      console.log(
        "BENEVOLE MIS A JOUR LIGNE",
        i+1
      );



      break;


    }


  }



  clearBenevolesCache();



  return true;


}









/* =====================================================
   CHANGEMENT STATUT SIMPLE
===================================================== */


function updateStatutBenevole(
  idBenevole,
  statut
){


  idBenevole =
    String(idBenevole)
    .padStart(3,"0");



  const sheet =
    SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      "INSCRIPTIONS"
    );



  const data =
    sheet.getDataRange().getValues();



  const headers =
    data[0].map(function(h){
      return String(h).trim();
    });



  const colId =
    headers.indexOf(
      "ID_BENEVOLE"
    );


  const colStatut =
    headers.indexOf(
      "STATUT"
    );



  for(let i=1;i<data.length;i++){


    if(

      String(data[i][colId])
      .padStart(3,"0")

      ===

      idBenevole

    ){


      sheet
        .getRange(
          i+1,
          colStatut+1
        )
        .setValue(
          statut
        );


      break;


    }


  }



  clearBenevolesCache();



  return true;


}







/* =====================================================
   TEST MANUEL
===================================================== */


function testValidation004(){

  console.log(
    "DEBUT TEST 004"
  );


  validerBenevoleComplet("004");


  console.log(
    "FIN TEST 004"
  );

}