/*
 * Faux Office.js, uniquement pour le banc d'essai hors Outlook.
 * Reproduit la partie de l'API que le volet utilise, avec une valeur
 * delayDeliveryTime gardee en memoire, et permet de simuler des echecs.
 *
 * Ce fichier n'est jamais servi au complement reel.
 */

(function () {
  "use strict";

  // Scenarios pilotables par l'URL : ?no113=1, ?echecSet=1, ?echecSave=1,
  // ?echecGet=1, ?delai=<millisecondes epoch>
  var p = new URLSearchParams(location.search);

  var etat = {
    delai: p.has("delai") ? Number(p.get("delai")) : 0,
    supporte113: !p.has("no113"),
    echecSet: p.has("echecSet") ? "Erreur simulee a l'ecriture." : null,
    echecGet: p.has("echecGet") ? "Erreur simulee a la lecture." : null,
    echecSave: p.has("echecSave") ? "Erreur simulee a l'enregistrement." : null,
    journal: []
  };

  function repondre(rappel, valeur, erreur) {
    setTimeout(function () {
      if (erreur) {
        rappel({ status: "failed", error: { message: erreur } });
      } else {
        rappel({ status: "succeeded", value: valeur });
      }
    }, 60);
  }

  window.Office = {
    HostType: { Outlook: "Outlook" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },

    onReady: function (rappel) {
      setTimeout(function () { rappel({ host: "Outlook" }); }, 10);
    },

    context: {
      requirements: {
        isSetSupported: function (nom, version) {
          if (nom === "Mailbox" && version === "1.13") { return etat.supporte113; }
          return true;
        }
      },
      mailbox: {
        item: {
          delayDeliveryTime: {
            getAsync: function (rappel) {
              etat.journal.push("getAsync -> " + etat.delai);
              repondre(rappel, etat.delai, etat.echecGet);
            },
            setAsync: function (valeur, rappel) {
              var ms = (valeur instanceof Date) ? valeur.getTime() : valeur;
              etat.journal.push("setAsync <- " + ms +
                (ms ? " (" + new Date(ms).toISOString() + ")" : " (efface)"));
              if (!etat.echecSet) { etat.delai = ms; }
              repondre(rappel, undefined, etat.echecSet);
            }
          },
          saveAsync: function (rappel) {
            etat.journal.push("saveAsync");
            repondre(rappel, "faux-id", etat.echecSave);
          }
        }
      }
    }
  };

  // Manette exposee au banc d'essai.
  window.__banc = etat;
})();
