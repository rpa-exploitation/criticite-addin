/* global Office */

// Compatible API Mailbox 1.5 (Exchange 2019 on-premises)
// Le niveau de criticité est porté par le préfixe d'objet [C1]-[C4]
// et mémorisé en customProperties (API 1.0).

const PROP_NAME = "criticite";
const PREFIX_REGEX = /^\[C[1-4]\]\s*/;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.querySelectorAll(".level").forEach((btn) => {
    btn.addEventListener("click", () => applyLevel(btn));
  });

  // Affiche le niveau déjà choisi si le volet est rouvert (lecture de l'objet)
  Office.context.mailbox.item.subject.getAsync((res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      const m = (res.value || "").match(/^\[(C[1-4])\]/);
      if (m) {
        const btn = document.querySelector(`.level[data-level="${m[1]}"]`);
        if (btn) markSelected(btn);
        setStatus(`Niveau actuel : ${m[1]}`, "ok");
      }
    }
  });
});

function applyLevel(btn) {
  const level = btn.dataset.level;
  const item = Office.context.mailbox.item;

  // 1. Préfixe dans l'objet — le marqueur principal (visible + exploitable par règles de transport)
  item.subject.getAsync((subjRes) => {
    if (subjRes.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Impossible de lire l'objet. Réessayez.", "err");
      return;
    }
    const cleanSubject = (subjRes.value || "").replace(PREFIX_REGEX, "");
    item.subject.setAsync(`[${level}] ${cleanSubject}`, (setRes) => {
      if (setRes.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Impossible de modifier l'objet. Réessayez.", "err");
        return;
      }

      // 2. Mémorisation en propriété personnalisée (bonus, API 1.0)
      item.loadCustomPropertiesAsync((propRes) => {
        if (propRes.status === Office.AsyncResultStatus.Succeeded) {
          const props = propRes.value;
          props.set(PROP_NAME, level);
          props.saveAsync(() => {
            markSelected(btn);
            setStatus(`Niveau ${level} appliqué. Vous pouvez envoyer.`, "ok");
          });
        } else {
          // L'objet est préfixé : c'est l'essentiel
          markSelected(btn);
          setStatus(`Niveau ${level} appliqué. Vous pouvez envoyer.`, "ok");
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
