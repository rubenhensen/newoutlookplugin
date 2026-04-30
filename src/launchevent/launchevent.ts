// OnMessageSend handler. Runs in a separate WebView runtime from the
// taskpane, so it cannot read in-memory taskpane state. It uses an x-
// prefixed internet header set by the taskpane (x-pg-encrypt-on-send)
// plus the attachment list to decide whether the message is allowed
// through. Custom properties were tried first but did not propagate
// cross-runtime in new Outlook (OWA-based).
//
// Behavior:
//  - encrypt-on-send not requested      → allow.
//  - requested + already encrypted       → allow.
//  - requested + not yet encrypted       → block, point user at the
//                                          PostGuard taskpane via commandId.

/* global Office */

const HEADER_ENCRYPT_ON_SEND = "x-pg-encrypt-on-send";
const POSTGUARD_ENCRYPTED_FILENAME = "postguard.encrypted";
const COMPOSE_BUTTON_ID = "postGuardComposeButton";

const BLOCKED_MESSAGE =
  "PostGuard is on but this message is not encrypted yet. " +
  "Open the PostGuard taskpane and click Encrypt & Send.";

function log(msg: string): void {
  // eslint-disable-next-line no-console
  console.log(`[pg-launchevent] ${msg}`);
}

// Belt-and-suspenders: if the handler ever takes more than 12 seconds, just
// allow the send. Better to fail open than have new Outlook hang forever
// behind the "PostGuard duurt langer dan verwacht" dialog.
function allowAfterTimeout(event: Office.AddinCommands.Event, ms = 12000): () => void {
  const timer = setTimeout(() => {
    log(`fallback timeout (${ms}ms) reached; allowing the send`);
    try {
      event.completed({ allowEvent: true });
    } catch (e) {
      log(`fallback event.completed threw: ${String(e)}`);
    }
  }, ms);
  return () => clearTimeout(timer);
}

function onMessageSendHandler(event: Office.AddinCommands.Event): void {
  log("onMessageSendHandler invoked");
  const cancelTimeout = allowAfterTimeout(event);

  const item = Office.context.mailbox.item as Office.MessageCompose;
  if (!item) {
    log("no mailbox item; allowing");
    cancelTimeout();
    event.completed({ allowEvent: true });
    return;
  }

  item.internetHeaders.getAsync([HEADER_ENCRYPT_ON_SEND], (hdrRes) => {
    log(`internetHeaders.getAsync status=${hdrRes.status}`);
    if (hdrRes.status !== Office.AsyncResultStatus.Succeeded) {
      cancelTimeout();
      event.completed({ allowEvent: true });
      return;
    }

    const encryptRequested = hdrRes.value[HEADER_ENCRYPT_ON_SEND] === "true";
    log(`encryptRequested=${encryptRequested}`);
    if (!encryptRequested) {
      cancelTimeout();
      event.completed({ allowEvent: true });
      return;
    }

    item.getAttachmentsAsync((attRes) => {
      log(`getAttachmentsAsync status=${attRes.status}`);
      const attachments = attRes.status === Office.AsyncResultStatus.Succeeded
        ? attRes.value
        : [];
      const alreadyEncrypted = attachments.some(
        (a) => a.name?.toLowerCase() === POSTGUARD_ENCRYPTED_FILENAME
      );
      log(`alreadyEncrypted=${alreadyEncrypted} (${attachments.length} attachments)`);

      cancelTimeout();
      if (alreadyEncrypted) {
        event.completed({ allowEvent: true });
        return;
      }

      const opts: Office.SmartAlertsEventCompletedOptions = {
        allowEvent: false,
        errorMessage: BLOCKED_MESSAGE,
        commandId: COMPOSE_BUTTON_ID,
      };
      event.completed(opts);
    });
  });
}

log("script loaded");
Office.onReady((info) => {
  log(`Office.onReady fired; host=${info?.host} platform=${info?.platform}`);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  log("handler associated");
});
