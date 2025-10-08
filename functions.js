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
    try {
      toast("Launching new email…");
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["support@abelwomack.com"],
        subject: "IT Support Request",
        htmlBody:
          "<p>Please describe your issue, impact, and urgency.</p>" +
          "<p><b>Device/User:</b><br/><b>Location:</b><br/><b>Apps affected:</b><br/><b>When it started:</b></p>"
      });
    } catch (e) {
      // Works even outside Outlook/OWA so you can test in a normal browser
      try { window.open("mailto:support@abelwomack.com?subject=IT%20Support%20Request", "_blank"); } catch (_) {}
    } finally { safeComplete(event); }
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
