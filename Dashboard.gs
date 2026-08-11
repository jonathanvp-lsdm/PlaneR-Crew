function getDashboardData() {

  const edition = Core_EditionService.getEditionActive();

  if (!edition) {

    throw new Error("Aucune édition active.");

  }

  const idEdition = edition.ID;

  const sheet = SpreadsheetApp
      .getActive()
      .getSheetByName("INSCRIPTIONS");

  const data = sheet.getDataRange().getValues();

  const headers = data.shift();

  const COL_ID_EDITION = headers.indexOf("ID_EDITION");
  const COL_STATUT     = headers.indexOf("STATUT");

  if (COL_ID_EDITION === -1 || COL_STATUT === -1) {

    throw new Error(
      "Colonnes ID_EDITION ou STATUT introuvables."
    );

  }

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

  /* =====================================================
     ALERTES DU DASHBOARD
  ===================================================== */

  const alertes = [];

  if (attente > 0) {

    alertes.push({

      niveau: "critical",
      texte: attente + " candidature(s) en attente de validation."

    });

  }

  if (valides === 0) {

    alertes.push({

      niveau: "warning",
      texte: "Aucun bénévole validé pour cette édition."

    });

  }

  if (inscriptions === 0) {

    alertes.push({

      niveau: "info",
      texte: "Aucune inscription enregistrée."

    });

  }

  /* =====================================================
     RETOUR
  ===================================================== */

  return {

      inscriptions,
      valides,
      attente,
      refuses,

      alertes

  };

}