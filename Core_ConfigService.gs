/* =====================================================
   PLANE'R CREW
   CONFIG SERVICE
===================================================== */

/* =====================================================
   LECTURE D'UN PARAMÈTRE
===================================================== */

function getConfig(cle){

  const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("CONFIG");

  if(!sheet){

    throw new Error(
      "Feuille CONFIG introuvable."
    );

  }

  const data = sheet
      .getDataRange()
      .getValues();

  for(let i = 1; i < data.length; i++){

    if(data[i][0] === cle){

      return data[i][1];

    }

  }

  return null;

}

/* =====================================================
   ÉCRITURE D'UN PARAMÈTRE
===================================================== */

function setConfig(cle, valeur){

  const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("CONFIG");

  if(!sheet){

    throw new Error(
      "Feuille CONFIG introuvable."
    );

  }

  const data = sheet
      .getDataRange()
      .getValues();

  for(let i = 1; i < data.length; i++){

    if(data[i][0] === cle){

      sheet
        .getRange(i + 1, 2)
        .setValue(valeur);

      return true;

    }

  }

  sheet.appendRow([

    cle,

    valeur,

    ""

  ]);

  return true;

}