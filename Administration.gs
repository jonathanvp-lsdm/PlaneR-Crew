/* =====================================================
   PLANE'R CREW
   MODULE ADMINISTRATION
   Version 2.0
===================================================== */

/* =====================================================
   CONFIGURATION CRM
===================================================== */

function getConfigurationCRM() {

  return {

    edition:
      Core_EditionService.getEditionActive(),

    inscriptions:
      getConfig("INSCRIPTIONS_OUVERTES"),

    developpement:
      Core_ConfigService.isModeDeveloppement(),

    version:
      CONFIG.VERSION

  };

}

/* =====================================================
   EDITIONS
===================================================== */

function getEditions() {

  return Core_EditionService
    .getToutesLesEditions();

}