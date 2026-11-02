/*
 * Copier pour Odoo - logique de mise en forme (sans dépendance à Office.js,
 * pour pouvoir être testée hors d'Outlook).
 */
(function (root) {
  "use strict";

  var SEPARATOR = "________________________________________";

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function pad2(n) {
    return (n < 10 ? "0" : "") + n;
  }

  // "Mercredi 09 septembre 2026 13:40" (même format que l'historique Outlook)
  function formatDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return "";
    var s = new Intl.DateTimeFormat("fr-BE", {
      weekday: "long",
      day: "2-digit",
      month: "long",
      year: "numeric"
    }).format(d);
    s = s.charAt(0).toUpperCase() + s.slice(1);
    return s + " " + pad2(d.getHours()) + ":" + pad2(d.getMinutes());
  }

  function formatAddress(a) {
    if (!a) return "";
    var name = (a.displayName || "").trim();
    var mail = (a.emailAddress || "").trim();
    if (name && mail && name.toLowerCase() !== mail.toLowerCase()) return name + " <" + mail + ">";
    return mail || name;
  }

  function formatAddressList(list) {
    return (list || []).map(formatAddress).filter(Boolean).join("; ");
  }

  // Lignes d'entête : [[libellé, valeur], ...]
  function headerLines(meta) {
    var lines = [
      ["De", formatAddress(meta.from)],
      ["Envoyé", formatDate(meta.date)],
      ["À", formatAddressList(meta.to)]
    ];
    var cc = formatAddressList(meta.cc);
    if (cc) lines.push(["Cc", cc]);
    lines.push(["Objet", meta.subject || ""]);
    return lines;
  }

  // --- Coupure de l'historique (messages cités sous le dernier message) ---

  var HISTORY_SELECTORS = [
    "#appendonsend",                       // Nouvel Outlook / OWA
    "#divRplyFwdMsg",                      // Nouvel Outlook / OWA
    "#mail-editor-reference-message-container",
    "div.gmail_quote",                     // Gmail
    "blockquote[type='cite']",             // Apple Mail / Thunderbird
    "div[style*='border-top:solid #E1E1E1']", // Outlook desktop
    "div[style*='border-top: solid #E1E1E1']",
    "div[style*='border-top:solid #B5C4DF']",
    "div[style*='border-top: solid #B5C4DF']",
    "div.OutlookMessageHeader"
  ];

  var HEADER_TEXT_RE = /^\s*(De|From|Van|Von)\s*:[\s\S]{0,400}?(Envoy[ée]|Sent|Date|Verzonden|Gesendet)\s*:/i;
  var ORIGINAL_RE = /^\s*-{2,}\s*(Original Message|Message d'origine|Oorspronkelijk bericht)\s*-{2,}/i;

  function isBefore(a, b) {
    // true si a précède b dans le document
    return !!(a.compareDocumentPosition(b) & 4);
  }

  function findHistoryStart(body) {
    var best = null;
    var bySel = body.querySelector(HISTORY_SELECTORS.join(","));
    if (bySel) best = bySel;

    var candidates = body.querySelectorAll("p, div, span, font");
    for (var i = 0; i < candidates.length; i++) {
      var el = candidates[i];
      if (best && !isBefore(el, best)) break; // on ne cherche que plus haut
      var txt = el.textContent || "";
      if (HEADER_TEXT_RE.test(txt) || ORIGINAL_RE.test(txt)) {
        best = el;
        break;
      }
    }
    return best;
  }

  function isEmptyNode(n) {
    if (n.nodeType === 3) return !/\S/.test(n.nodeValue.replace(/ /g, " "));
    if (n.nodeType === 8) return true;
    if (n.nodeType !== 1) return true;
    var tag = n.tagName.toLowerCase();
    if (tag === "br" || tag === "hr") return true;
    if (tag === "img" || tag === "table") return false;
    if (n.querySelector("img, table")) return false;
    return !/\S/.test((n.textContent || "").replace(/ /g, " "));
  }

  function trimTrailing(el) {
    while (el.lastChild && isEmptyNode(el.lastChild)) el.removeChild(el.lastChild);
    if (el.lastChild && el.lastChild.nodeType === 1) trimTrailing(el.lastChild);
  }

  function cutFrom(start, body) {
    var n = start;
    while (n && n !== body) {
      while (n.nextSibling) n.parentNode.removeChild(n.nextSibling);
      n = n.parentNode;
    }
    start.parentNode.removeChild(start);
    trimTrailing(body);
  }

  // Nettoie le HTML du mail. options.onlyLast : retire l'historique ; options.noImages : retire les images.
  function cleanBody(doc, options) {
    var body = doc.body;
    var junk = body.querySelectorAll("script, style, meta, link, title, xml, iframe, frame, object, embed, form, input, button, textarea, select, svg, math, base");
    for (var i = 0; i < junk.length; i++) junk[i].parentNode && junk[i].parentNode.removeChild(junk[i]);

    // Sécurité : aucun gestionnaire d'événement ni URL javascript: ne doit survivre
    var all = body.querySelectorAll("*");
    for (var k = 0; k < all.length; k++) {
      var attrs = all[k].attributes;
      for (var a = attrs.length - 1; a >= 0; a--) {
        var name = attrs[a].name.toLowerCase();
        var val = (attrs[a].value || "").replace(/[\s\u0000-\u001f]/g, "").toLowerCase();
        if (name.indexOf("on") === 0 || name === "srcdoc" || name === "formaction" ||
            (/^(href|src|action|background|xlink:href|data)$/.test(name) && /^(javascript|vbscript|data:text)/.test(val))) {
          all[k].removeAttribute(attrs[a].name);
        }
      }
    }

    if (options.onlyLast) {
      var start = findHistoryStart(body);
      if (start && start !== body) {
        var backup = body.innerHTML;
        cutFrom(start, body);
        if (!/\S/.test((body.textContent || "").replace(/ /g, " "))) body.innerHTML = backup; // sécurité
      }
    }

    var imgs = body.querySelectorAll("img");
    for (var j = 0; j < imgs.length; j++) {
      var src = imgs[j].getAttribute("src") || "";
      if (options.noImages || /^cid:/i.test(src)) imgs[j].parentNode.removeChild(imgs[j]);
    }
    trimTrailing(body);
    return body;
  }

  // Conversion simple HTML -> texte
  function toText(node) {
    var out = "";
    var BLOCK = /^(p|div|h[1-6]|li|tr|table|blockquote|ul|ol|pre)$/i;
    function children(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) walk(c);
    }
    function walk(n) {
      if (n.nodeType === 3) { out += n.nodeValue.replace(/\s+/g, " "); return; }
      if (n.nodeType !== 1) return;
      var tag = n.tagName.toLowerCase();
      if (tag === "br") { out += "\n"; return; }
      if (tag === "td" || tag === "th") { children(n); out += "\t"; return; }
      var block = BLOCK.test(tag);
      if (block && out && !/\n$/.test(out)) out += "\n";
      children(n);
      if (block && !/\n$/.test(out)) out += "\n";
    }
    children(node);
    return out
      .replace(/ /g, " ")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  /*
   * meta : { from, to[], cc[], subject, date }
   * bodyHtml : HTML complet du mail
   * domImpl : document.implementation (pour créer un document de travail isolé)
   * options : { onlyLast: bool, noImages: bool, separator: bool }
   * -> { html, text }
   */
  function build(meta, bodyHtml, domImpl, options) {
    options = options || {};
    var holder = domImpl.createHTMLDocument("");
    holder.documentElement.innerHTML = /<body[\s>]/i.test(bodyHtml || "")
      ? bodyHtml.replace(/^[\s\S]*?<html[^>]*>/i, "").replace(/<\/html>[\s\S]*$/i, "")
      : "<body>" + (bodyHtml || "") + "</body>";
    var body = cleanBody(holder, options);

    var lines = headerLines(meta);
    var sep = options.separator === false ? "" : SEPARATOR;

    var headerHtml =
      "<p>" + (sep ? sep + "<br>" : "") +
      lines.map(function (l) { return "<b>" + l[0] + ":</b> " + escapeHtml(l[1]); }).join("<br>") +
      "</p>";
    var html = headerHtml + "<div>" + body.innerHTML + "</div>";

    var text =
      (sep ? sep + "\n" : "") +
      lines.map(function (l) { return l[0] + ": " + l[1]; }).join("\n") +
      "\n\n" + toText(body);

    return { html: html, text: text };
  }

  var api = {
    build: build,
    formatDate: formatDate,
    formatAddress: formatAddress,
    formatAddressList: formatAddressList,
    cleanBody: cleanBody,
    toText: toText
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.CopierOdoo = api;
})(this);
