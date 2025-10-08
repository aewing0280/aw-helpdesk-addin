// functions.js — command handlers for the AW Helpdesk add-in
(function () {
  'use strict';

  function safeComplete(event) {
    try { event && typeof event.completed === "function" && event.completed(); } catch (e) {}
  }

  function toast(msg) {
    try {
      var item = Office?.context?.mailbox?.item;
      item?.notificationMessages?.replaceAsync("aw-status", {
        type: "informationalMessage", message: msg, icon: "icon16", persistent: false
      });
    } catch (e) {}
  }

function createTicket(event) {
  const to = "support@abelwomack.com";
  const subject = "IT Support Request";
  const bodyText =
    "Please describe your issue, impact, and urgency.\n\n" +
    "Device/User:\nLocation:\nApps affected:\nWhen it started:\n";

  // Deep link to OWA compose (works in any browser if pop-ups allowed)
  const deeplink =
    `https://outlook.office.com/mail/deeplink/compose?` +
    `to=${encodeURIComponent(to)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyText)}`;

  // Same-origin redirect page on your GitHub Pages (avoids cross-domain quirks)
  const redirect =
    `https://aewing0280.github.io/aw-helpdesk-addin/compose.html?` +
    `to=${encodeURIComponent(to)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyText)}`;

  try {
    // 1) Prefer the Outlook API (no popup blockers)
    if (Office?.context?.mailbox?.displayNewMessageForm) {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [to],
        subject,
        htmlBody: "<pre style='font-family:Segoe UI,system-ui'>" + bodyText + "</pre>"
      });
      return;
    }

    // 2) Try opening OWA compose directly
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(deeplink);
      return;
    }

    // 3) Open same-origin redirect (then it forwards to OWA)
    window.open(redirect, "_blank");
  } finally {
    try { event?.completed?.(); } catch (_) {}
  }
}



function openPortal(event) {
  const url = "https://help.abelwomack.com";
  try {
    // Prefer the Office API if available (respects user gesture from ribbon click)
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
    } else {
      // Fallback—open a new tab/window
      window.open(url, "_blank");
    }
  } finally {
    try { event?.completed?.(); } catch (_) {}
  }
}


  // Expose globally for ExecuteFunction
  window.createTicket = createTicket;
  window.openPortal   = openPortal;

  if (window.Office && Office.onReady) { Office.onReady(function(){ /* ready */ }); }

  // Make testing obvious in the browser console
  console.log("AW functions.js loaded");
})();
