// Entry point for the PostGuard taskpane. Detects whether we are in compose
// or read mode and dispatches to the corresponding view.

import { isComposeMode } from "../lib/office-helpers";
import { mountComposeView } from "./compose-view";
import { mountReadView } from "./read-view";

/* global Office */

// Debug logging that surfaces in the taskpane itself. Used to diagnose
// hangs/errors without DevTools attached. Remove once stable.
function dlog(msg: string): void {
  const ts = new Date().toISOString().slice(11, 23);
  console.log(`[pg ${ts}] ${msg}`);
  const host = document.getElementById("view-loading");
  if (!host) return;
  let pre = document.getElementById("pg-debug-log") as HTMLPreElement | null;
  if (!pre) {
    pre = document.createElement("pre");
    pre.id = "pg-debug-log";
    pre.style.cssText =
      "font-family:Consolas,monospace;font-size:11px;text-align:left;white-space:pre-wrap;padding:8px;background:#f4f4f4;border:1px solid #ddd;margin-top:12px;max-height:300px;overflow:auto;";
    host.appendChild(pre);
  }
  pre.textContent += `[${ts}] ${msg}\n`;
}

window.addEventListener("error", (e) => {
  dlog(`window.error: ${e.message} @ ${e.filename}:${e.lineno}:${e.colno}`);
});
window.addEventListener("unhandledrejection", (e) => {
  const reason = e.reason instanceof Error ? `${e.reason.message}\n${e.reason.stack}` : String(e.reason);
  dlog(`unhandledrejection: ${reason}`);
});

dlog("script loaded; waiting for Office.onReady");

// Fail loudly if Office.onReady doesn't fire — points at office.js load failure.
const onReadyTimeout = window.setTimeout(() => {
  dlog("TIMEOUT: Office.onReady did not fire within 8s");
}, 8000);

const views = {
  loading: byId("view-loading"),
  compose: byId("view-compose"),
  read_encrypted: byId("view-read-encrypted"),
  read_was_encrypted: byId("view-read-was-encrypted"),
  read_noop: byId("view-read-noop"),
  decrypted: byId("view-decrypted"),
  policy_editor: byId("view-policy-editor"),
  yivi: byId("view-yivi"),
  error: byId("view-error"),
};

export type ViewName = keyof typeof views;

export function showView(name: ViewName): void {
  for (const [k, el] of Object.entries(views)) {
    if (el) el.hidden = k !== name;
  }
}

export function showError(message: string): void {
  const errEl = byId("pg-error-text");
  if (errEl) errEl.textContent = message;
  showView("error");
}

export function setStatus(message: string, kind: "info" | "error" = "info"): void {
  const el = byId("pg-status");
  if (!el) return;
  if (!message) {
    el.classList.add("pg-status-hidden");
    el.textContent = "";
    return;
  }
  el.classList.remove("pg-status-hidden");
  el.classList.toggle("pg-status-error", kind === "error");
  el.textContent = message;
}

function byId(id: string): HTMLElement | null {
  return document.getElementById(id);
}

Office.onReady((info) => {
  window.clearTimeout(onReadyTimeout);
  dlog(`Office.onReady fired; host=${info.host} platform=${info.platform}`);
  if (info.host !== Office.HostType.Outlook) {
    showError("PostGuard only runs inside Outlook.");
    return;
  }

  const retry = byId("pg-error-retry") as HTMLButtonElement | null;
  if (retry) retry.addEventListener("click", () => bootstrap());

  bootstrap();
});

async function bootstrap(): Promise<void> {
  showView("loading");
  setStatus("");
  dlog("bootstrap: start");
  try {
    const compose = isComposeMode();
    dlog(`bootstrap: isComposeMode=${compose}`);
    if (compose) {
      dlog("bootstrap: mountComposeView…");
      await mountComposeView();
      dlog("bootstrap: mountComposeView done");
    } else {
      dlog("bootstrap: mountReadView…");
      await mountReadView();
      dlog("bootstrap: mountReadView done");
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : "PostGuard failed to start.";
    const stack = err instanceof Error && err.stack ? err.stack : "";
    dlog(`bootstrap: ERROR ${message}\n${stack}`);
    showError(message);
  }
}
