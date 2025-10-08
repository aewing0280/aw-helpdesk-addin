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
    var url = "https://help.abelwomack.com";
    try {
      // Some portals block iframes; we try dialog first then fall back
      toast("Opening Helpdesk Portal…");
      Office.context.ui.displayDialogAsync(url, { height: 55, width: 40, displayInIframe: true }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { try { window.open(url, "_blank"); } catch(_){} }
      });
    } catch (e) {
      try { window.open(url, "_blank"); } catch(_) {}
    } finally { safeComplete(event); }
  }

  // Expose globally for ExecuteFunction
  window.createTicket = createTicket;
  window.openPortal   = openPortal;

  if (window.Office && Office.onReady) { Office.onReady(function(){ /* ready */ }); }

  // Make testing obvious in the browser console
  console.log("AW functions.js loaded");
})();
