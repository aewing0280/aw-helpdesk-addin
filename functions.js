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
  const body = encodeURIComponent(
    "Please describe your issue, impact, and urgency.\n\n" +
    "Device/User:\nLocation:\nApps affected:\nWhen it started:\n"
  );

  // OWA compose deep link: works reliably in browser and desktop (opens OWA)
  const deeplink =
    `https://outlook.office.com/mail/deeplink/compose?` +
    `to=${encodeURIComponent(to)}&subject=${encodeURIComponent(subject)}&body=${body}`;

  try {
    // Prefer sanctioned API if available to avoid popup blockers
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(deeplink);
    } else if (Office?.context?.mailbox?.displayNewMessageForm) {
      // Desktop Outlook often supports this
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [to],
        subject: subject,
        htmlBody: "<pre style='font-family:Segoe UI,system-ui'>" + decodeURIComponent(body) + "</pre>"
      });
    } else {
      // Last-resort fallback
      window.open(`mailto:${to}?subject=${encodeURIComponent(subject)}`, "_blank");
    }
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
