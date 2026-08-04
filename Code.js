/* =====================================================
   PLANE'R CREW - CODE PRINCIPAL
   VERSION 3.0.0 STABLE
===================================================== */


/* =====================================================
   APPLICATION PRINCIPALE
===================================================== */

function doGet() {

  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("Plane'R Crew")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}

function testInclude() {

  const txt = include("javascript");

  Logger.log(txt.substring(0,500));

}

/* =====================================================
   INCLUSION HTML
===================================================== */

function include(filename) {

  return HtmlService
    .createTemplateFromFile(filename)
    .evaluate()
    .getContent();

}


/* =====================================================
   CHARGEMENT DES PAGES
===================================================== */

function loadPage(page){

  try{

    return HtmlService
      .createTemplateFromFile(page)
      .evaluate()
      .getContent();

  }

  catch(err){

    Logger.log(err);

    return `
      <div class="alert alert-danger m-4">
        <h4>Erreur de chargement</h4>
        <p>${err.message}</p>
      </div>
    `;

  }

}


/* =====================================================
   TEST CONNEXION JAVASCRIPT
===================================================== */

function getTest(){

  return [

    {

      ID:"TEST",

      PRENOM:"Jonathan",

      NOM:"TEST",

      EMAIL:"test@test.com",

      STATUT:"EN ATTENTE",

      CHOIX_1:"BAR"

    }

  ];

}


/* =====================================================
   TEST INCLUDE JAVASCRIPT
===================================================== */

function testIncludeJavascript(){

  const contenu = include("javascript");

  Logger.log(contenu.substring(0,200));

}


/* =====================================================
   TEST DEPLOIEMENT
===================================================== */

function testDeploiement(){

  Logger.log("VERSION ACTIVE : 3.0.0 STABLE");

  Logger.log("DATE : " + new Date());

}


/* =====================================================
   TEST PING
===================================================== */

function ping(){

  return "PONG";

}