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

  // Who’s asking (best effort)
  const user = {
    name:  Office?.context?.mailbox?.userProfile?.displayName || "",
    email: Office?.context?.mailbox?.userProfile?.emailAddress || ""
  };
  const when = new Date().toLocaleString();

  // ----- Ticket template (text + html) -----
  const bodyText =
`Summary: <one sentence of the issue>

Impact/Urgency: <P1 | P2 | P3 | P4>
Users Affected: <who/which team>
Location: <office/remote + city>
Device: <model / asset tag>
Applications/Services Affected: <app names>
Error Messages: <exact text or screenshot>
When It Started: <date/time>
Steps Already Tried: <bullets>
Attachments: <logs/screenshots if any>
Best Callback: <phone or Teams handle>

Requested by: ${user.name} ${user.email ? `(${user.email})` : ""}
Timestamp: ${when}
`;

  const bodyHtml =
`<div style="font-family:Segoe UI,system-ui,Arial,sans-serif;font-size:12.5pt;line-height:1.35">
  <p><b>Summary:</b> &lt;one sentence of the issue&gt;</p>
  <p><b>Impact/Urgency:</b> &lt;P1 | P2 | P3 | P4&gt;</p>
  <p><b>Users Affected:</b> &lt;who/which team&gt;</p>
  <p><b>Location:</b> &lt;office/remote + city&gt;</p>
  <p><b>Device:</b> &lt;model / asset tag&gt;</p>
  <p><b>Applications/Services Affected:</b> &lt;app names&gt;</p>
  <p><b>Error Messages:</b> &lt;exact text or screenshot&gt;</p>
  <p><b>When It Started:</b> &lt;date/time&gt;</p>
  <p><b>Steps Already Tried:</b><br>
     • &lt;step 1&gt;<br>
     • &lt;step 2&gt;</p>
  <p><b>Attachments:</b> &lt;logs/screenshots if any&gt;</p>
  <p><b>Best Callback:</b> &lt;phone or Teams handle&gt;</p>
  <hr style="border:none;border-top:1px solid #ddd;margin:14px 0">
  <p style="color:#555"><b>Requested by:</b> ${user.name}${user.email ? ` (${user.email})` : ""}<br>
     <b>Timestamp:</b> ${when}</p>
</div>`;

  const deeplink =
    `https://outlook.office.com/mail/deeplink/compose?` +
    `to=${encodeURIComponent(to)}` +
    `&subject=${encodeURIComponent(subject)}` +
    `&body=${encodeURIComponent(bodyText)}`;

  const redirect =
    `https://aewing0280.github.io/aw-helpdesk-addin/compose.html?` +
    `to=${encodeURIComponent(to)}` +
    `&subject=${encodeURIComponent(subject)}` +
    `&body=${encodeURIComponent(bodyText)}`;

  try {
    // #1: Native API (preferred)
    if (Office?.context?.mailbox?.displayNewMessageForm) {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [to],
        subject,
        htmlBody: bodyHtml
        // Optional: cc the requester
        // ,ccRecipients: [ user.email ].filter(Boolean)
      });
      return;
    }

    // #2: Open OWA compose deeplink
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(deeplink);
      return;
    }

    // #3: Same-origin redirect → OWA compose
    window.open(redirect, "_blank");
  } finally {
    try { event?.completed?.(); } catch (_) {}
  }
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
