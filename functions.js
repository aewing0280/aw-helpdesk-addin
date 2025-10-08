
// functions.js — command handlers for the AW Helpdesk add-in
(function () {
  'use strict';

  function safeComplete(event) {
    try { event && typeof event.completed === "function" && event.completed(); } catch (e) {}
  }

  function toast(msg) {
    try {
      var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
      if (item && item.notificationMessages) {
        item.notificationMessages.replaceAsync("aw-status", {
          type: "informationalMessage",
          message: msg,
          icon: "icon16",
          persistent: false
        });
      }
    } catch (e) {}
  }

  // Opens a new pre-addressed message to support@abelwomack.com
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
      // Fallback outside Office
      try {
        window.open("mailto:support@abelwomack.com?subject=IT%20Support%20Support%20Request", "_blank");
      } catch (_) {}
    } finally {
      safeComplete(event);
    }
  }

  // Opens the helpdesk portal
  function openPortal(event) {
    var url = "https://help.abelwomack.com";
    try {
      toast("Opening Helpdesk Portal…");
      Office.context.ui.displayDialogAsync(url, { height: 55, width: 40, displayInIframe: true }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          try { window.open(url, "_blank"); } catch (_) {}
        }
      });
    } catch (e) {
      try { window.open(url, "_blank"); } catch (_) {}
    } finally {
      safeComplete(event);
    }
  }

  // Expose globally for ExecuteFunction
  window.createTicket = createTicket;
  window.openPortal   = openPortal;

  if (window.Office && Office.onReady) {
    Office.onReady(function(){ /* ready */ });
  }
})();
