function getDashboardData() {

  const edition = Core_EditionService.getEditionActive();

  if (!edition) {

    throw new Error("Aucune édition active.");

  }

  Logger.log(JSON.stringify(edition));

  const idEdition = edition.ID;

  const sheet = SpreadsheetApp
      .getActive()
      .getSheetByName("INSCRIPTIONS");

  const data = sheet.getDataRange().getValues();

  const headers = data.shift();

  const COL_ID_EDITION = headers.indexOf("ID_EDITION");
  const COL_STATUT     = headers.indexOf("STATUT");

  let inscriptions = 0;
  let valides = 0;
  let attente = 0;
  let refuses = 0;

  data.forEach(function(row){

      if(String(row[COL_ID_EDITION]) !== String(idEdition)) return;

      inscriptions++;

      switch(String(row[COL_STATUT]).toUpperCase()){

          case "VALIDE":
              valides++;
              break;

          case "ATTENTE":
              attente++;
              break;

          case "REFUSE":
              refuses++;
              break;

      }

  });

  return {

      inscriptions,
      valides,
      attente,
      refuses

  };

}