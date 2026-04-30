// OnMessageSend handler. Runs in a separate WebView runtime from the
// taskpane, so it cannot read in-memory taskpane state. It uses x-
// prefixed internet headers set by the taskpane plus the attachment
// list to decide whether the message is allowed through. Custom
// properties were tried first but did not propagate cross-runtime in
// new Outlook (OWA-based).
//
// Behavior:
//  - encrypt-on-send not requested            → allow.
//  - requested + encrypted + recipients match → allow.
//  - requested + not yet encrypted            → block (encrypt prompt).
//  - requested + encrypted + recipients drift → block (re-encrypt prompt).

/* global Office */

const HEADER_ENCRYPT_ON_SEND = "x-pg-encrypt-on-send";
const HEADER_ENCRYPTED_RECIPIENTS = "x-pg-encrypted-recipients";
const POSTGUARD_ENCRYPTED_FILENAME = "postguard.encrypted";
const COMPOSE_BUTTON_ID = "postGuardComposeButton";

const NOT_ENCRYPTED_MESSAGE =
  "PostGuard is on but this message is not encrypted yet. " +
  "Open the PostGuard taskpane and click Encrypt & Send.";

const STALE_ENCRYPTION_MESSAGE =
  "PostGuard recipients or settings changed since the last encryption. " +
  "Open the PostGuard taskpane and click Re-encrypt & Send before sending.";

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

function block(event: Office.AddinCommands.Event, errorMessage: string): void {
  const opts: Office.SmartAlertsEventCompletedOptions = {
    allowEvent: false,
    errorMessage,
    commandId: COMPOSE_BUTTON_ID,
  };
  event.completed(opts);
}

function recipientsKey(addresses: Office.EmailAddressDetails[]): string {
  return addresses
    .map((a) => (a.emailAddress ?? "").toLowerCase().trim())
    .filter(Boolean)
    .sort()
    .join(",");
}

function getRecipientsAsync(
  recipients: Office.Recipients
): Promise<Office.EmailAddressDetails[]> {
  return new Promise((resolve) => {
    recipients.getAsync((res) =>
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : [])
    );
  });
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

  item.internetHeaders.getAsync(
    [HEADER_ENCRYPT_ON_SEND, HEADER_ENCRYPTED_RECIPIENTS],
    (hdrRes) => {
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

      const stampedRecipients = hdrRes.value[HEADER_ENCRYPTED_RECIPIENTS] ?? "";

      item.getAttachmentsAsync(async (attRes) => {
        log(`getAttachmentsAsync status=${attRes.status}`);
        const attachments =
          attRes.status === Office.AsyncResultStatus.Succeeded ? attRes.value : [];
        const alreadyEncrypted = attachments.some(
          (a) => a.name?.toLowerCase() === POSTGUARD_ENCRYPTED_FILENAME
        );
        log(`alreadyEncrypted=${alreadyEncrypted} (${attachments.length} attachments)`);

        if (!alreadyEncrypted) {
          cancelTimeout();
          block(event, NOT_ENCRYPTED_MESSAGE);
          return;
        }

        // Verify the encryption matches the message's current To+Cc list. The
        // taskpane clears HEADER_ENCRYPTED_RECIPIENTS on policy/sign drift; we
        // also re-derive the recipient list here so a recipient added behind
        // the taskpane's back is still caught.
        const [to, cc] = await Promise.all([
          getRecipientsAsync(item.to),
          getRecipientsAsync(item.cc),
        ]);
        const currentKey = recipientsKey([...to, ...cc]);
        const stale = stampedRecipients === "" || currentKey !== stampedRecipients;
        log(`stamped=${stampedRecipients || "<empty>"} current=${currentKey} stale=${stale}`);

        cancelTimeout();
        if (stale) {
          block(event, STALE_ENCRYPTION_MESSAGE);
          return;
        }
        event.completed({ allowEvent: true });
      });
    }
  );
}

log("script loaded");
Office.onReady((info) => {
  log(`Office.onReady fired; host=${info?.host} platform=${info?.platform}`);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  log("handler associated");
});
