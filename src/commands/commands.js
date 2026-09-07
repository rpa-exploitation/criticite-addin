/* global Office */

const HEADER_NAME = "X-Criticite";
const VALID_LEVELS = ["C1", "C2", "C3", "C4"];

/**
 * Handler de l'événement ItemSend (on-send).
 * Autorise l'envoi uniquement si le header X-Criticite est présent et valide.
 * @param {Office.AddinCommands.Event} event
 */
function validateCriticality(event) {
  const item = Office.context.mailbox.item;

  item.internetHeaders.getAsync([HEADER_NAME], (res) => {
    const value =
      res.status === Office.AsyncResultStatus.Succeeded ? res.value[HEADER_NAME] : null;

    if (value && VALID_LEVELS.includes(value)) {
      // Criticité valide : l'envoi continue
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
          "\u00ab Niveau de criticité \u00bb du ruban.",
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
