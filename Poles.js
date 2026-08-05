/* =====================================================
   RECUPERATION DES POLES
===================================================== */

function getPoles(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("POLES");

  if(!sheet){
    throw new Error("Feuille POLES introuvable.");
  }

  const data =
    sheet.getDataRange().getValues();

  const headers = data[0];

  const result = [];

  for(let i=1;i<data.length;i++){

    const ligne = {};

    headers.forEach(function(header,index){

      ligne[header] = data[i][index];

    });

    if(String(ligne.ACTIF).toUpperCase() === "OUI"){

      result.push(ligne);

    }

  }

  result.sort(function(a,b){

    return Number(a.ORDRE) - Number(b.ORDRE);

  });

  return result;

}

/* =====================================================
   TEST
===================================================== */

function testGetPoles(){

  Logger.log(
    JSON.stringify(
      getPoles(),
      null,
      2
    )
  );

}