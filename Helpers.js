/* ==========================================================
   CHARGEMENT DES PAGES HTML DYNAMIQUES
========================================================== */


function loadPage(page){


  return HtmlService

    .createTemplateFromFile(page)

    .evaluate()

    .getContent();


}



/* ==========================================================
   INCLUSION DES FICHIERS HTML
========================================================== */


function include(filename){


  return HtmlService

    .createHtmlOutputFromFile(filename)

    .getContent();


}