/* global Office */

// Compatible API Mailbox 1.5 (Exchange 2019 on-premises) — syntaxe ES5
// Version de production : préfixe d'objet [C1]-[C4] + étiquette centrée
// colorée en tête du corps. (Le module de sensibilité native via EWS a été
// retiré : la plateforme hébergée renvoie une réponse vide aux requêtes EWS
// des compléments — réactivable si l'opérateur débloque EWS.)

var PROP_NAME = "criticite";
var PREFIX_REGEX = /^\[C[1-4]\]\s*/;
var TAG_START = "<!--CRITICITE-TAG-START-->";
var TAG_END = "<!--CRITICITE-TAG-END-->";
var TAG_REGEX = /<!--CRITICITE-TAG-START-->[\s\S]*?<!--CRITICITE-TAG-END-->/g;
// OWA supprime les commentaires HTML : reconnaissance de secours par le contenu
var TAG_FALLBACK_REGEX = /<div[^>]*>\s*<span[^>]*>\s*C[1-4]\s*-\s*(?:Strictement confidentiel|Confidentiel|Interne|Public)\s*<\/span>\s*<\/div>/g;

var LEVELS = {
  C1: { label: "Strictement confidentiel", color: "#7f1d1d", bg: "#fde8e8", border: "#dc2626", sensitivity: "Confidential" },
  C2: { label: "Confidentiel",             color: "#7c2d12", bg: "#fff3cd", border: "#ea580c", sensitivity: "Confidential" },
  C3: { label: "Interne",                  color: "#1e3a8a", bg: "#e8f0fe", border: "#2563eb", sensitivity: "Confidential" },
  C4: { label: "Public",                   color: "#14532d", bg: "#e7f6ec", border: "#16a34a", sensitivity: "Normal" }
};

function buildTag(level) {
  var l = LEVELS[level];
  return TAG_START +
    '<div style="text-align:center;margin:0 0 10px 0;">' +
    '<span style="display:inline-block;background:' + l.bg +
    ';border:1px solid ' + l.border + ';color:' + l.color +
    ';padding:2px 10px;font-family:Segoe UI,Arial,sans-serif;font-size:12px;' +
    'font-weight:600;border-radius:3px;">' +
    level + ' - ' + l.label +
    '</span></div>' + TAG_END;
}

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;

  var buttons = document.querySelectorAll(".level");
  for (var i = 0; i < buttons.length; i++) {
    (function (btn) {
      btn.addEventListener("click", function () { applyLevel(btn); });
    })(buttons[i]);
  }

  Office.context.mailbox.item.subject.getAsync(function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      var m = (res.value || "").match(/^\[(C[1-4])\]/);
      if (m) {
        var btn = document.querySelector('.level[data-level="' + m[1] + '"]');
        if (btn) markSelected(btn);
        setStatus("Niveau actuel : " + m[1], "ok");
      }
    }
  });
});

function applyLevel(btn) {
  var level = btn.getAttribute("data-level");
  var item = Office.context.mailbox.item;

  setStatus("Application du niveau " + level + "\u2026", "");

  // 1. Préfixe dans l'objet — marqueur principal
  item.subject.getAsync(function (subjRes) {
    if (subjRes.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Impossible de lire l'objet. R\u00e9essayez.", "err");
      return;
    }
    var cleanSubject = (subjRes.value || "").replace(PREFIX_REGEX, "");
    item.subject.setAsync("[" + level + "] " + cleanSubject, function (setRes) {
      if (setRes.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Impossible de modifier l'objet. R\u00e9essayez.", "err");
        return;
      }

      // 2. Étiquette dans le corps (remplacement si déjà présente)
      item.body.getAsync(Office.CoercionType.Html, function (bodyRes) {
        if (bodyRes.status !== Office.AsyncResultStatus.Succeeded) {
          finish(btn, level, "Objet pr\u00e9fix\u00e9, mais corps illisible (\u00e9tiquette non ins\u00e9r\u00e9e).");
          return;
        }

        var html = bodyRes.value || "";
        var tag = buildTag(level);

        // Supprimer TOUTES les étiquettes existantes (avec ou sans marqueurs,
        // OWA supprimant les commentaires HTML), puis insérer la nouvelle.
        var cleaned = html.replace(TAG_REGEX, "").replace(TAG_FALLBACK_REGEX, "");

        var newHtml;
        var bodyTag = cleaned.match(/<body[^>]*>/i);
        if (bodyTag) {
          var idx = cleaned.indexOf(bodyTag[0]) + bodyTag[0].length;
          newHtml = cleaned.slice(0, idx) + tag + cleaned.slice(idx);
        } else {
          newHtml = tag + cleaned;
        }

        item.body.setAsync(newHtml, { coercionType: Office.CoercionType.Html }, function (bodySetRes) {
          if (bodySetRes.status === Office.AsyncResultStatus.Succeeded) {
            finish(btn, level, null);
          } else {
            finish(btn, level, "Objet pr\u00e9fix\u00e9, mais \u00e9tiquette non ins\u00e9r\u00e9e.");
          }
        });
      });
    });
  });
}

function finish(btn, level, warning) {
  try {
    Office.context.mailbox.item.loadCustomPropertiesAsync(function (propRes) {
      if (propRes.status === Office.AsyncResultStatus.Succeeded) {
        var props = propRes.value;
        props.set(PROP_NAME, level);
        props.saveAsync(function () {});
      }
    });
  } catch (e) { /* non bloquant */ }

  markSelected(btn);
  if (warning) {
    setStatus(warning, "err");
  } else {
    setStatus("Niveau " + level + " appliqu\u00e9. Vous pouvez envoyer.", "ok");
  }
}

function markSelected(btn) {
  var buttons = document.querySelectorAll(".level");
  for (var i = 0; i < buttons.length; i++) {
    buttons[i].classList.remove("selected");
    buttons[i].setAttribute("aria-checked", "false");
  }
  btn.classList.add("selected");
  btn.setAttribute("aria-checked", "true");
}

function setStatus(msg, cls) {
  var el = document.getElementById("status");
  el.textContent = msg;
  el.className = cls || "";
}
