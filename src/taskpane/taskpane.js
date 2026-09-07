/* global Office */

// Compatible API Mailbox 1.5 (Exchange 2019 on-premises) — syntaxe ES5
// v3 : en plus du préfixe d'objet [C1]-[C4], insère un bandeau HTML coloré
// en tête du corps du message. Le bandeau est balisé par des marqueurs
// pour être remplacé proprement si l'utilisateur change de niveau.

var PROP_NAME = "criticite";
var PREFIX_REGEX = /^\[C[1-4]\]\s*/;
var BANNER_START = "<!--CRITICITE-BANNER-START-->";
var BANNER_END = "<!--CRITICITE-BANNER-END-->";
// Supprime un bandeau existant (avec son conteneur éventuel)
var BANNER_REGEX = /<!--CRITICITE-BANNER-START-->[\s\S]*?<!--CRITICITE-BANNER-END-->/g;

var LEVELS = {
  C1: {
    label: "CRITIQUE",
    text: "Ce message requiert une action imm\u00e9diate.",
    color: "#7f1d1d", border: "#dc2626", bg: "#fde8e8"
  },
  C2: {
    label: "IMPORTANT",
    text: "Un traitement prioritaire est attendu.",
    color: "#7c2d12", border: "#ea580c", bg: "#fff3e8"
  },
  C3: {
    label: "NORMAL",
    text: "Flux de travail habituel.",
    color: "#1e3a8a", border: "#2563eb", bg: "#e8f0fe"
  },
  C4: {
    label: "INFORMATION",
    text: "Aucune action requise.",
    color: "#374151", border: "#6b7280", bg: "#f3f4f6"
  }
};

function buildBanner(level) {
  var l = LEVELS[level];
  return BANNER_START +
    '<div style="background:' + l.bg + ';border-left:6px solid ' + l.border +
    ';padding:10px 14px;margin:0 0 14px 0;font-family:Segoe UI,Arial,sans-serif;' +
    'font-size:13px;color:' + l.color + ';">' +
    '<b>CRITICIT\u00c9 ' + level + ' \u2014 ' + l.label + '</b><br/>' +
    l.text +
    '</div>' + BANNER_END;
}

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;

  var buttons = document.querySelectorAll(".level");
  for (var i = 0; i < buttons.length; i++) {
    (function (btn) {
      btn.addEventListener("click", function () { applyLevel(btn); });
    })(buttons[i]);
  }

  // Affiche le niveau déjà choisi si le volet est rouvert (lecture de l'objet)
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

      // 2. Bandeau dans le corps (remplacement si déjà présent)
      item.body.getAsync(Office.CoercionType.Html, function (bodyRes) {
        if (bodyRes.status !== Office.AsyncResultStatus.Succeeded) {
          finish(btn, level, "Objet pr\u00e9fix\u00e9, mais corps illisible (bandeau non ins\u00e9r\u00e9).");
          return;
        }

        var html = bodyRes.value || "";
        var banner = buildBanner(level);
        var newHtml;

        if (BANNER_REGEX.test(html)) {
          // Un bandeau existe : on le remplace sur place
          newHtml = html.replace(BANNER_REGEX, banner);
        } else {
          // Pas de bandeau : insertion en tête du corps.
          // On insère juste après <body...> si présent, sinon en préfixe brut.
          var bodyTag = html.match(/<body[^>]*>/i);
          if (bodyTag) {
            var idx = html.indexOf(bodyTag[0]) + bodyTag[0].length;
            newHtml = html.slice(0, idx) + banner + html.slice(idx);
          } else {
            newHtml = banner + html;
          }
        }

        item.body.setAsync(newHtml, { coercionType: Office.CoercionType.Html }, function (bodySetRes) {
          if (bodySetRes.status === Office.AsyncResultStatus.Succeeded) {
            finish(btn, level, null);
          } else {
            finish(btn, level, "Objet pr\u00e9fix\u00e9, mais bandeau non ins\u00e9r\u00e9.");
          }
        });
      });
    });
  });
}

function finish(btn, level, warning) {
  // 3. Mémorisation en propriété personnalisée (bonus, silencieux)
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
    setStatus("Niveau " + level + " appliqu\u00e9 (objet + bandeau). Vous pouvez envoyer.", "ok");
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
