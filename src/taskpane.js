/*
 * Envoi programmé côté serveur pour Outlook classique.
 *
 * Utilise item.delayDeliveryTime (requirement set Mailbox 1.13), dont le délai est
 * traité par Exchange et non par le client : le message part à l'heure prévue même
 * si Outlook est fermé et le poste éteint. À ne pas confondre avec l'option native
 * « Ne pas envoyer avant », qui laisse le message dans la Boîte d'envoi locale.
 */

"use strict";

var VERSION = "1.0.0";
var HORIZON_JOURS = 30;      // borne que l'on s'impose, Exchange n'en documente aucune
var MARGE_MINIMALE_MS = 60 * 1000;

var el = {};
var item = null;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) { return; }

  ["chargement", "incompatible", "incompatible-detail", "app", "statut", "statut-texte",
   "rapides", "date", "heure", "programmer", "annuler", "message", "version"]
    .forEach(function (id) { el[id] = document.getElementById(id); });

  el.version.textContent = "v" + VERSION;

  if (!Office.context.requirements.isSetSupported("Mailbox", "1.13")) {
    el["incompatible-detail"].textContent =
      "Diagnostic : Mailbox 1.13 non disponible sur ce client.";
    montrer("incompatible");
    return;
  }

  item = Office.context.mailbox.item;
  if (!item || !item.delayDeliveryTime) {
    el["incompatible-detail"].textContent =
      "Diagnostic : la propriété delayDeliveryTime est absente de cet élément.";
    montrer("incompatible");
    return;
  }

  construireRapides();
  el.programmer.addEventListener("click", programmerDepuisChamps);
  el.annuler.addEventListener("click", retirerProgrammation);

  var defaut = prochaineHeureRonde();
  el.date.value = versValeurDate(defaut);
  el.heure.value = versValeurHeure(defaut);

  rafraichirStatut(function () { montrer("app"); });
});

/* ---------- Lecture de l'état courant ---------- */

function rafraichirStatut(apres) {
  item.delayDeliveryTime.getAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      afficherMessage("Impossible de lire la programmation : " + res.error.message, "erreur");
      if (apres) { apres(); }
      return;
    }

    if (!res.value || res.value === 0) {
      el.statut.className = "statut statut-aucun";
      el["statut-texte"].textContent =
        "Ce message partira dès que vous cliquerez sur Envoyer.";
      el.annuler.classList.add("masque");
    } else {
      var d = new Date(res.value);
      el.statut.className = "statut statut-programme";
      el["statut-texte"].innerHTML =
        "Envoi programmé le <strong>" + formatLong(d) + "</strong>.<br>" +
        "Cliquez sur Envoyer pour confirmer.";
      el.annuler.classList.remove("masque");
      el.date.value = versValeurDate(d);
      el.heure.value = versValeurHeure(d);
    }
    if (apres) { apres(); }
  });
}

/* ---------- Écriture ---------- */

function programmerDepuisChamps() {
  if (!el.date.value || !el.heure.value) {
    afficherMessage("Indiquez une date et une heure.", "erreur");
    return;
  }
  var d = el.date.value.split("-");
  var h = el.heure.value.split(":");
  programmer(new Date(+d[0], +d[1] - 1, +d[2], +h[0], +h[1], 0, 0));
}

function programmer(quand) {
  var erreur = valider(quand);
  if (erreur) { afficherMessage(erreur, "erreur"); return; }

  basculerBoutons(false);
  item.delayDeliveryTime.setAsync(quand, function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      basculerBoutons(true);
      afficherMessage("Échec de la programmation : " + res.error.message, "erreur");
      return;
    }
    // On enregistre le brouillon pour que le réglage survive à une fermeture de la fenêtre.
    item.saveAsync(function (sauve) {
      basculerBoutons(true);
      if (sauve.status !== Office.AsyncResultStatus.Succeeded) {
        afficherMessage("Envoi programmé le " + formatLong(quand) +
          ". Le brouillon n'a pas pu être enregistré, évitez de fermer la fenêtre.", "succes");
      } else {
        afficherMessage("Envoi programmé le " + formatLong(quand) + ".", "succes");
      }
      rafraichirStatut();
    });
  });
}

function retirerProgrammation() {
  basculerBoutons(false);
  item.delayDeliveryTime.setAsync(0, function (res) {
    basculerBoutons(true);
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      afficherMessage("Impossible de retirer la programmation : " + res.error.message, "erreur");
      return;
    }
    afficherMessage("Programmation retirée. Le message partira immédiatement à l'envoi.", "succes");
    rafraichirStatut();
  });
}

function valider(quand) {
  if (!(quand instanceof Date) || isNaN(quand.getTime())) {
    return "Date ou heure invalide.";
  }
  if (quand.getTime() <= Date.now() + MARGE_MINIMALE_MS) {
    return "Choisissez un moment au moins une minute dans le futur.";
  }
  var limite = Date.now() + HORIZON_JOURS * 24 * 60 * 60 * 1000;
  if (quand.getTime() > limite) {
    return "Au-delà de " + HORIZON_JOURS + " jours, le comportement n'est pas garanti. " +
           "Choisissez une date plus rapprochée.";
  }
  return null;
}

/* ---------- Choix rapides ---------- */

function construireRapides() {
  choixRapides().forEach(function (choix) {
    var b = document.createElement("button");
    b.type = "button";
    b.innerHTML = "<strong>" + choix.libelle + "</strong><span class='quand'>" +
                  formatCourt(choix.date) + "</span>";
    b.addEventListener("click", function () { programmer(choix.date); });
    el.rapides.appendChild(b);
  });
}

function choixRapides() {
  var liste = [];

  var dansUneHeure = new Date(Date.now() + 60 * 60 * 1000);
  dansUneHeure.setSeconds(0, 0);
  liste.push({ libelle: "Dans 1 heure", date: dansUneHeure });

  var demain = aJour(1);
  liste.push({ libelle: "Demain 7 h 30", date: aHeure(demain, 7, 30) });
  liste.push({ libelle: "Demain 13 h 00", date: aHeure(demain, 13, 0) });

  // Prochain lundi, ou le lundi suivant si on est déjà lundi. On l'écarte quand
  // il tombe le même jour que « Demain », pour ne pas proposer deux fois la
  // même chose un dimanche.
  var lundi = new Date();
  var delta = (8 - lundi.getDay()) % 7 || 7;
  lundi.setDate(lundi.getDate() + delta);
  lundi = aHeure(lundi, 7, 30);
  if (lundi.toDateString() !== demain.toDateString()) {
    liste.push({ libelle: "Lundi 7 h 30", date: lundi });
  }

  return liste;
}

function aJour(n) {
  var d = new Date();
  d.setDate(d.getDate() + n);
  return d;
}

function aHeure(d, h, m) {
  var r = new Date(d.getTime());
  r.setHours(h, m, 0, 0);
  return r;
}

function prochaineHeureRonde() {
  var d = new Date(Date.now() + 60 * 60 * 1000);
  d.setMinutes(0, 0, 0);
  return d;
}

/* ---------- Présentation ---------- */

function formatLong(d) {
  return d.toLocaleDateString("fr-CA", {
    weekday: "long", day: "numeric", month: "long"
  }) + " à " + d.toLocaleTimeString("fr-CA", { hour: "2-digit", minute: "2-digit" });
}

function formatCourt(d) {
  return d.toLocaleDateString("fr-CA", { day: "numeric", month: "short" }) +
         " " + d.toLocaleTimeString("fr-CA", { hour: "2-digit", minute: "2-digit" });
}

function versValeurDate(d) {
  return d.getFullYear() + "-" + deuxChiffres(d.getMonth() + 1) + "-" + deuxChiffres(d.getDate());
}

function versValeurHeure(d) {
  return deuxChiffres(d.getHours()) + ":" + deuxChiffres(d.getMinutes());
}

function deuxChiffres(n) { return (n < 10 ? "0" : "") + n; }

function montrer(id) {
  ["chargement", "incompatible", "app"].forEach(function (autre) {
    el[autre].classList.toggle("masque", autre !== id);
  });
}

function afficherMessage(texte, genre) {
  el.message.textContent = texte;
  el.message.className = "message message-" + genre;
}

function basculerBoutons(actif) {
  el.programmer.disabled = !actif;
  el.annuler.disabled = !actif;
  Array.prototype.forEach.call(el.rapides.querySelectorAll("button"), function (b) {
    b.disabled = !actif;
  });
}
