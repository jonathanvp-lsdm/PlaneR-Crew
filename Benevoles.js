/* =====================================================
   PLANE'R CREW
   RECUPERATION BENEVOLES + INSCRIPTIONS
   AVEC CACHE SERVEUR
===================================================== */


/*
   VERSION NETTOYEE
   - Conservation ID bénévole
   - Conservation ID inscription
   - Suppression doublon validation
*/


function getBenevoles(){

  const cache = CacheService.getScriptCache();

  const cached = cache.get("BENEVOLES_CACHE");


  if(cached){

    console.log(
      "DONNEES SERVIES DEPUIS LE CACHE"
    );

    return JSON.parse(cached);

  }



  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  const sheetBenevoles =
    ss.getSheetByName("BENEVOLES");


  const sheetInscriptions =
    ss.getSheetByName("INSCRIPTIONS");



  if(!sheetBenevoles || !sheetInscriptions){

    throw new Error(
      "Feuille BENEVOLES ou INSCRIPTIONS introuvable"
    );

  }



  const benevolesData =
    sheetBenevoles.getDataRange().getValues();


  const inscriptionsData =
    sheetInscriptions.getDataRange().getValues();



  const benevolesHeaders =
    benevolesData[0];


  const inscriptionsHeaders =
    inscriptionsData[0];



  let benevoles = {};



  for(let i=1;i<benevolesData.length;i++){


    let obj={};



    benevolesHeaders.forEach(function(header,index){


      let value =
        benevolesData[i][index];



      if(value instanceof Date){

        value =
          Utilities.formatDate(
            value,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          );

      }



      obj[header]=value;


    });



    benevoles[obj.ID]=obj;


  }





  let result=[];



  for(let i=1;i<inscriptionsData.length;i++){


    let inscription={};



    inscriptionsHeaders.forEach(function(header,index){


      let value =
        inscriptionsData[i][index];



      if(value instanceof Date){

        value =
          Utilities.formatDate(
            value,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          );

      }



      inscription[header]=value;


    });



    const fiche =
      benevoles[inscription.ID_BENEVOLE] || {};



    result.push({

      ...fiche,

      ...inscription,


      ID:
        fiche.ID,


      ID_INSCRIPTION:
        inscription.ID

    });



  }




  cache.put(

    "BENEVOLES_CACHE",

    JSON.stringify(result),

    300

  );



  console.log(
    "DONNEES MISES EN CACHE :",
    result.length
  );



  return result;


}







function clearBenevolesCache(){


  CacheService
    .getScriptCache()
    .remove(
      "BENEVOLES_CACHE"
    );



  console.log(
    "CACHE BENEVOLES SUPPRIME"
  );


}







function updateStatutBenevole(
  idInscription,
  nouveauStatut
){


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  const sheet =
    ss.getSheetByName("INSCRIPTIONS");



  const data =
    sheet.getDataRange().getValues();



  const headers =
    data[0];



  const colId =
    headers.indexOf("ID");


  const colStatut =
    headers.indexOf("STATUT");



  for(let i=1;i<data.length;i++){


    if(

      String(data[i][colId])
      ===
      String(idInscription)

    ){


      sheet
      .getRange(
        i+1,
        colStatut+1
      )
      .setValue(
        nouveauStatut
      );



      break;

    }


  }



  clearBenevolesCache();



  return true;


}