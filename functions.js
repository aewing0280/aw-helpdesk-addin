// Keep everything on window and always call event.completed()
(function () {
  function isCompose() {
    // Compose has setAsync methods; read doesn't
    return !!(Office?.context?.mailbox?.item?.to?.setAsync);
  }

  function log(msg, err) {
    try { console.log("[AW]", msg, err || ""); } catch(_){}
    // Try to surface an in-product notice (best-effort, safe to ignore on read surface)
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "awdiag",
        { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: String(msg).slice(0,150),
          icon: "icon16",
          persistent: false
        },
        function(){}
      );
    } catch(_){}
  }

  function buildBodies() {
    const user = Office?.context?.mailbox?.userProfile || {};
    const when = new Date().toLocaleString();
    const text = `Summary: <one sentence>
Impact/Urgency: <P1|P2|P3|P4>
Users Affected: <>
Location: <>
Device: <>
Apps/Services: <>
Error Messages: <>
Start Time: <>
Steps Tried: <>
Callback: <>

Requested by: ${user.displayName || ""} ${user.emailAddress ? "(" + user.emailAddress + ")" : ""}
Timestamp: ${when}
`;
    const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:12.5pt;line-height:1.35">
  <p><b>Summary:</b> &lt;one sentence&gt;</p>
  <p><b>Impact/Urgency:</b> &lt;P1|P2|P3|P4&gt;</p>
  <p><b>Users Affected:</b> &lt;&gt;</p>
  <p><b>Location:</b> &lt;&gt;</p>
  <p><b>Device:</b> &lt;&gt;</p>
  <p><b>Apps/Services:</b> &lt;&gt;</p>
  <p><b>Error Messages:</b> &lt;&gt;</p>
  <p><b>Start Time:</b> &lt;&gt;</p>
  <p><b>Steps Tried:</b><br>• &lt;&gt;<br>• &lt;&gt;</p>
  <p><b>Callback:</b> &lt;phone or Teams handle&gt;</p>
</div>`;
    return { text, html };
  }

  // === REQUIRED: expose handlers on window and call event.completed() ===

  window.createTicket = function (event) {
    const to = "support@abelwomack.com";
    const subject = "IT Support Request";
    const { text, html } = buildBodies();

    try {
      if (isCompose()) {
        // Fill current draft
        const item = Office.context.mailbox.item;
        const ops = [];
        ops.push(new Promise(res => item.to.setAsync([{ emailAddress: to }], {}, () => res())));
        ops.push(new Promise(res => item.subject.setAsync(subject, {}, () => res())));
        ops.push(new Promise(res => item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, () => res())));
        Promise.all(ops).then(() => log("compose: filled")).finally(() => event.completed());
      } else {
        // Read surface: open new compose (best effort)
        if (Office?.context?.mailbox?.displayNewMessageForm) {
          Office.context.mailbox.displayNewMessageForm({ toRecipients: [to], subject, htmlBody: html });
          log("read: displayNewMessageForm");
        } else {
          const deeplink =
            `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(to)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(text)}`;
          if (Office?.context?.ui?.openBrowserWindow) {
            Office.context.ui.openBrowserWindow(deeplink);
            log("read: openBrowserWindow");
          } else {
            window.open(deeplink, "_blank");
            log("read: window.open");
          }
        }
        event.completed();
      }
    } catch (e) {
      log("createTicket error", e);
      try { event.completed(); } catch(_) {}
    }
  };

  window.openPortal = function (event) {
    try {
      const url = "https://help.abelwomack.com";
      if (Office?.context?.ui?.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(url);
        log("portal: openBrowserWindow");
      } else {
        window.open(url, "_blank");
        log("portal: window.open");
      }
    } catch (e) {
      log("openPortal error", e);
    } finally {
      try { event.completed(); } catch(_) {}
    }
  };
})();
