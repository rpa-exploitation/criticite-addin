/* global Office */

// Compatible API Mailbox 1.5 (Exchange 2019 on-premises)
// Validation on-send : l'objet doit commencer par [C1]-[C4].
// v3 : Office.initialize explicite (requis par Outlook desktop pour les
// pages UI-less) + fail-open pour ne jamais bloquer l'envoi sur erreur.

var SUBJECT_REGEX = /^\[C[1-4]\]/;
var TIMEOUT_MS = 10000;

// Requis : sans cette déclaration, Outlook classique peut ne jamais
// finaliser l'exécution du handler (symptôme : "traite votre demande" en boucle).
Office.initialize = function () {};

/**
 * Handler de l'événement ItemSend (on-send).
 * Autorise l'envoi uniquement si l'objet est préfixé d'un niveau valide.
 * @param {Office.AddinCommands.Event} event
 */
function validateCriticality(event) {
  var done = false;

  function complete(allow) {
    if (done) return;
    done = true;
    event.completed({ allowEvent: allow });
  }

  // Filet de sécurité : si rien n'a répondu en 10 s, on laisse partir le mail
  // plutôt que de bloquer l'utilisateur indéfiniment.
  setTimeout(function () { complete(true); }, TIMEOUT_MS);

  try {
    var item = Office.context.mailbox.item;

    item.subject.getAsync(function (res) {
      var subject =
        res.status === Office.AsyncResultStatus.Succeeded ? res.value || "" : "";

      if (SUBJECT_REGEX.test(subject)) {
        // Criticité présente : l'envoi continue
        complete(true);
        return;
      }

      // Criticité absente : on bloque et on guide l'utilisateur
      try {
        item.notificationMessages.replaceAsync(
          "criticiteManquante",
          {
            type: "errorMessage",
            message:
              "Envoi bloqué : choisissez un niveau de criticité via le bouton " +
              "\u00ab Niveau de criticité \u00bb.",
          },
          function () {
            complete(false);
          }
        );
      } catch (e) {
        // La notification a échoué : on bloque quand même, message ou pas
        complete(false);
      }
    });
  } catch (e) {
    // Erreur imprévue : fail-open
    complete(true);
  }
}

// Exposition globale requise par le FunctionFile
window.validateCriticality = validateCriticality;
if (Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("validateCriticality", validateCriticality);
}
