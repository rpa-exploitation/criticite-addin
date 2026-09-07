/* global Office */

// Compatible API Mailbox 1.5 (Exchange 2019 on-premises)
// Validation on-send : l'objet doit commencer par [C1]-[C4].

const SUBJECT_REGEX = /^\[C[1-4]\]/;

/**
 * Handler de l'événement ItemSend (on-send).
 * Autorise l'envoi uniquement si l'objet est préfixé d'un niveau valide.
 * @param {Office.AddinCommands.Event} event
 */
function validateCriticality(event) {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((res) => {
    const subject = res.status === Office.AsyncResultStatus.Succeeded ? res.value || "" : "";

    if (SUBJECT_REGEX.test(subject)) {
      // Criticité présente : l'envoi continue
      event.completed({ allowEvent: true });
      return;
    }

    // Criticité absente : on bloque et on guide l'utilisateur
    item.notificationMessages.replaceAsync(
      "criticiteManquante",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message:
          "Envoi bloqué : choisissez un niveau de criticité via le bouton " +
          "\u00ab Niveau de criticité \u00bb.",
      },
      () => {
        event.completed({ allowEvent: false });
      }
    );
  });
}

// Exposition globale requise par le FunctionFile
window.validateCriticality = validateCriticality;
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("validateCriticality", validateCriticality);
}
