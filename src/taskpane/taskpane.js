/* global Office */

const HEADER_NAME = "X-Criticite";
const PREFIX_REGEX = /^\[C[1-4]\]\s*/;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.querySelectorAll(".level").forEach((btn) => {
    btn.addEventListener("click", () => applyLevel(btn));
  });

  // Affiche le niveau déjà choisi si le volet est rouvert
  Office.context.mailbox.item.internetHeaders.getAsync([HEADER_NAME], (res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded && res.value[HEADER_NAME]) {
      const current = res.value[HEADER_NAME];
      const btn = document.querySelector(`.level[data-level="${current}"]`);
      if (btn) markSelected(btn);
      setStatus(`Niveau actuel : ${current}`, "ok");
    }
  });
});

function applyLevel(btn) {
  const level = btn.dataset.level;
  const item = Office.context.mailbox.item;

  // 1. Header SMTP custom (exploitable par règles de transport / SIEM)
  const headers = {};
  headers[HEADER_NAME] = level;

  item.internetHeaders.setAsync(headers, (headerRes) => {
    if (headerRes.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Erreur lors de l'écriture du header. Réessayez.", "err");
      return;
    }

    // 2. Préfixe dans l'objet (visible par tous les destinataires)
    item.subject.getAsync((subjRes) => {
      if (subjRes.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Header écrit, mais objet illisible.", "err");
        return;
      }
      const cleanSubject = (subjRes.value || "").replace(PREFIX_REGEX, "");
      item.subject.setAsync(`[${level}] ${cleanSubject}`, (setRes) => {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          markSelected(btn);
          setStatus(`Niveau ${level} appliqué. Vous pouvez envoyer.`, "ok");
        } else {
          setStatus("Header écrit, mais préfixe d'objet non appliqué.", "err");
        }
      });
    });
  });
}

function markSelected(btn) {
  document.querySelectorAll(".level").forEach((b) => {
    b.classList.remove("selected");
    b.setAttribute("aria-checked", "false");
  });
  btn.classList.add("selected");
  btn.setAttribute("aria-checked", "true");
}

function setStatus(msg, cls) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = cls || "";
}
