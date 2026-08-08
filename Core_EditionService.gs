/* =====================================================
   PLANE'R CREW
   CORE EDITION SERVICE
   Version 1.1
===================================================== */

const Core_EditionService = (() => {

  const SHEET_NAME = "EDITIONS";

  /* =====================================================
     FEUILLE
  ===================================================== */

  function getSheet() {

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error(
        "La feuille EDITIONS est introuvable."
      );
    }

    return sheet;

  }

/* =====================================================
   LECTURE DES DONNEES
===================================================== */

function getData() {

  const sheet = getSheet();

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values.shift();

  const colSupprime = headers.indexOf("SUPPRIME");

  return values

    .filter(row => {

      // Si la colonne SUPPRIME n'existe pas,
      // on retourne toutes les lignes.
      if (colSupprime === -1) {
        return true;
      }

      return row[colSupprime] !== true &&
       row[colSupprime] !== "TRUE";

    })

    .map(row => {

      const edition = {};

      headers.forEach((header, index) => {

        let value = row[index];

        // Conversion des dates pour google.script.run
        if (value instanceof Date) {

          value = Utilities.formatDate(
            value,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          );

        }

        edition[header] = value;

      });

      return edition;

    });

}

  /* =====================================================
     EDITION ACTIVE
  ===================================================== */

  function getEditionActive() {

    return getData().find(
      edition => edition.EST_ACTIVE === true
    ) || null;

  }

  /* =====================================================
     LISTE DES EDITIONS
  ===================================================== */

  function getEditions() {

    return getData();

  }

  /* =====================================================
     RECHERCHE PAR ID
  ===================================================== */

  function getEditionById(id) {

    return getData().find(
      edition => Number(edition.ID) === Number(id)
    ) || null;

  }

  /* =====================================================
     DEFINIR L'EDITION ACTIVE
  ===================================================== */

  function setEditionActive(id) {

    const sheet = getSheet();

    const values = sheet.getDataRange().getValues();

    const headers = values.shift();

    const colId = headers.indexOf("ID");

    const colActive = headers.indexOf("EST_ACTIVE");

    if (colId === -1 || colActive === -1) {

      throw new Error(
        "Colonnes ID ou EST_ACTIVE introuvables."
      );

    }

    values.forEach((row, index) => {

      const estActive =
        Number(row[colId]) === Number(id);

      sheet
        .getRange(index + 2, colActive + 1)
        .setValue(estActive);

    });

  }

  /* =====================================================
     MODIFIER L'ETAT
  ===================================================== */

  function setEtatEdition(id, etat) {

    const sheet = getSheet();

    const values = sheet.getDataRange().getValues();

    const headers = values.shift();

    const colId = headers.indexOf("ID");

    const colEtat = headers.indexOf("ETAT");

    if (colId === -1 || colEtat === -1) {

      throw new Error(
        "Colonne ETAT introuvable."
      );

    }

    values.forEach((row, index) => {

      if (Number(row[colId]) === Number(id)) {

        sheet
          .getRange(index + 2, colEtat + 1)
          .setValue(etat);

      }

    });

  }

/* =====================================================
   CREER UNE EDITION
===================================================== */

function createEdition(data) {

    Logger.log("===== CREATE =====");
    Logger.log(JSON.stringify(data));

  const sheet = getSheet();

  const headers =
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
         .getValues()[0];

  const row = headers.map(function(header) {

    if (header === "ID") {

      const editions = getData();

      const maxId = editions.length
        ? Math.max(...editions.map(e => Number(e.ID) || 0))
        : 0;

      return maxId + 1;

    }

    if (header === "SUPPRIME") {
    return false;
}

    return data[header] !== undefined
    ? data[header]
    : "";

  });

sheet.appendRow(row);

const newId = row[headers.indexOf("ID")];

if (data.EST_ACTIVE === true) {
    setEditionActive(newId);
}

return true;

}

/* =====================================================
   MODIFIER UNE EDITION
===================================================== */

function updateEdition(data) {

    Logger.log("===== UPDATE =====");
    Logger.log(JSON.stringify(data));

  const sheet = getSheet();

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  const colId = headers.indexOf("ID");

  if (colId === -1) {
    throw new Error("Colonne ID introuvable.");
  }

  values.forEach((row, index) => {

    if (Number(row[colId]) === Number(data.ID)) {

      headers.forEach((header, colIndex) => {

        if (header === "ID") {
          return;
        }

        sheet
          .getRange(index + 2, colIndex + 1)
          .setValue(data[header]);

      });

    }

  });

  if (data.EST_ACTIVE === true) {
    setEditionActive(data.ID);
  }

  return true;

}

/* =====================================================
   SUPPRIMER UNE EDITION
===================================================== */

function deleteEdition(id) {

    const edition = getEditionById(id);

    if (!edition) {
        throw new Error("Édition introuvable.");
    }

   if (String(edition.EST_ACTIVE) === "true" ||
    String(edition.EST_ACTIVE) === "TRUE") {

    throw new Error("Impossible de supprimer l'édition active.");
    }

    const sheet = getSheet();

    const values = sheet.getDataRange().getValues();

    const headers = values.shift();

    const colId = headers.indexOf("ID");
    const colSupprime = headers.indexOf("SUPPRIME");

    if (colId === -1 || colSupprime === -1) {
        throw new Error("Colonne SUPPRIME introuvable.");
    }

    values.forEach((row, index) => {

        if (Number(row[colId]) === Number(id)) {

            sheet
                .getRange(index + 2, colSupprime + 1)
                .setValue(true);

        }

    });

    return true;

}

  /* =====================================================
     API PUBLIQUE
  ===================================================== */

return {

  getEditionActive,
  getEditions,
  getEditionById,

  createEdition,
  updateEdition,
  deleteEdition,

  setEditionActive,
  setEtatEdition

};

})();

/* =====================================================
   TESTS
===================================================== */

function testEditionActive() {

  Logger.log(
    Core_EditionService.getEditionActive()
  );

}

function testGetEditionById() {

  Logger.log(
    Core_EditionService.getEditionById(1)
  );

}

function testSetEtatEdition() {

  Core_EditionService.setEtatEdition(
    1,
    "INSCRIPTIONS_OUVERTES"
  );

  Logger.log(
    Core_EditionService.getEditionById(1)
  );

}

function testGetToutesLesEditions() {

  const editions =
    Core_EditionService.getToutesLesEditions();

  Logger.log(
    JSON.stringify(editions, null, 2)
  );

}

/* =====================================================
   API GOOGLE SCRIPT
===================================================== */

function updateEdition(data){

  return Core_EditionService.updateEdition(data);

}

function createEdition(data){

  return Core_EditionService.createEdition(data);

}

function deleteEdition(id){
    return Core_EditionService.deleteEdition(id);
}