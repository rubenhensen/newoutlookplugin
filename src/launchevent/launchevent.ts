// Event-based activation handler for new Outlook on Windows.
//
// Even though our manifest does not declare a <LaunchEvent>, new Outlook
// (and some tenant configurations) probe `launchevent.html?et=…` on Send
// and wait ~15s for an event.completed() callback before showing a Smart
// Alerts "PostGuard timed out" dialog. Serving this file with a passthrough
// handler unblocks Send instantly.
//
// All handlers must allow the event because compose-view.ts already calls
// saveItem() before yielding control to the user, so the draft is fully
// committed by the time Send fires.

/* global Office */

const allowEvent = (event: Office.AddinCommands.Event): void => {
  event.completed({ allowEvent: true });
};

Office.onReady(() => {
  // Register every handler name new Outlook might dispatch to us. The
  // Office runtime resolves `et=…` to one of these via the manifest, but
  // when the manifest doesn't list it Outlook falls back to common names.
  Office.actions.associate("onMessageSendHandler", allowEvent);
  Office.actions.associate("onAppointmentSendHandler", allowEvent);
  Office.actions.associate("onNewMessageComposeHandler", allowEvent);
  Office.actions.associate("onNewAppointmentComposeHandler", allowEvent);
  Office.actions.associate("onMessageAttachmentsChangedHandler", allowEvent);
  Office.actions.associate("onMessageRecipientsChangedHandler", allowEvent);
});
