// Inline policy editor used both for "Manage Access" (recipient policies)
// and "Sign" (sender attribute selection). Renders into #view-policy-editor.
//
// UI pattern mirrors postguard-website's RecipientSelectionFields: the email
// is locked at the top, selected attributes are rendered as input rows with
// a × delete button, and unselected attributes are shown as "+" chips that
// add a fresh row when clicked.

import {
  AttributeDescriptor,
  EMAIL_ATTRIBUTE_TYPE,
  SUPPORTED_ATTRIBUTES,
} from "../lib/attributes";
import { Policy, AttributeRequest } from "../lib/types";
import { t } from "../lib/i18n";
import { showView } from "./taskpane";

interface PolicyEditorOptions {
  initialPolicy: Policy;
  sign: boolean;
  onSave: (next: Policy) => void;
  onCancel: () => void;
}

// Working state for the currently open editor session. Email is implicit
// (always emitted on save), so this map holds the *extra* attributes only.
let workingPolicy = new Map<string, AttributeRequest[]>();

export function openPolicyEditor(opts: PolicyEditorOptions): void {
  const titleEl = document.getElementById("pg-policy-title")!;
  const recipientsEl = document.getElementById("pg-policy-recipients")!;
  const saveBtn = document.getElementById("pg-policy-save") as HTMLButtonElement;
  const cancelBtn = document.getElementById("pg-policy-cancel") as HTMLButtonElement;

  titleEl.textContent = opts.sign ? t("policyEditorTitleSign") : t("policyEditorTitle");
  saveBtn.textContent = t("policyEditorSave");
  cancelBtn.textContent = t("policyEditorCancel");

  workingPolicy = new Map();
  for (const [email, attrs] of Object.entries(opts.initialPolicy)) {
    const extras = attrs
      .filter((a) => a.t !== EMAIL_ATTRIBUTE_TYPE)
      .map((a) => ({ t: a.t, v: a.v }));
    workingPolicy.set(email, extras);
  }

  recipientsEl.innerHTML = "";
  for (const email of workingPolicy.keys()) {
    recipientsEl.appendChild(renderRecipient(email));
  }

  // Replace listeners by cloning the buttons (cheap way to drop prior bindings).
  const saveClone = saveBtn.cloneNode(true) as HTMLButtonElement;
  saveBtn.replaceWith(saveClone);
  saveClone.addEventListener("click", () => opts.onSave(collect()));

  const cancelClone = cancelBtn.cloneNode(true) as HTMLButtonElement;
  cancelBtn.replaceWith(cancelClone);
  cancelClone.addEventListener("click", () => opts.onCancel());

  showView("policy_editor");
}

function renderRecipient(email: string): HTMLElement {
  const section = document.createElement("div");
  section.className = "pg-policy-recipient";
  section.dataset.email = email;
  rerenderRecipient(section, email);
  return section;
}

function rerenderRecipient(section: HTMLElement, email: string): void {
  const extras = workingPolicy.get(email)!;
  section.innerHTML = "";

  const heading = document.createElement("div");
  heading.className = "pg-policy-recipient-email";
  heading.textContent = email;
  section.appendChild(heading);

  for (let i = 0; i < extras.length; i++) {
    const desc = SUPPORTED_ATTRIBUTES.find((d) => d.type === extras[i].t);
    if (!desc) continue;
    section.appendChild(
      renderAttrRow(extras, i, desc, () => rerenderRecipient(section, email))
    );
  }

  const addable = SUPPORTED_ATTRIBUTES.filter(
    (d) => d.type !== EMAIL_ATTRIBUTE_TYPE && !extras.some((e) => e.t === d.type)
  );
  if (addable.length > 0) {
    const addRow = document.createElement("div");
    addRow.className = "pg-policy-add-row";
    for (const desc of addable) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "pg-policy-add-chip";
      btn.textContent = `+ ${t(desc.type, desc.defaultLabel)}`;
      btn.addEventListener("click", () => {
        extras.push({ t: desc.type, v: "" });
        rerenderRecipient(section, email);
      });
      addRow.appendChild(btn);
    }
    section.appendChild(addRow);
  }
}

function renderAttrRow(
  extras: AttributeRequest[],
  index: number,
  desc: AttributeDescriptor,
  onDelete: () => void
): HTMLElement {
  const attr = extras[index];

  const row = document.createElement("div");
  row.className = "pg-policy-attr";

  const label = document.createElement("label");
  label.textContent = t(desc.type, desc.defaultLabel);
  row.appendChild(label);

  const inputRow = document.createElement("div");
  inputRow.className = "pg-policy-attr-input";

  const input = document.createElement("input");

  if (desc.type === "pbdf.gemeente.personalData.dateofbirth") {
    // Yivi stores DOB as DD-MM-YYYY but <input type="date"> uses YYYY-MM-DD.
    // Round-trip through helpers so the IBE identity at encrypt matches what
    // Yivi discloses at decrypt.
    input.type = "date";
    input.value = ddmmyyyyToHtml(attr.v);
    input.addEventListener("input", () => {
      attr.v = htmlToDdmmyyyy(input.value);
    });
  } else if (desc.type === "pbdf.sidn-pbdf.mobilenumber.mobilenumber") {
    // Yivi stores numbers in E.164. We don't have libphonenumber here yet so
    // we accept whatever the user types; mismatched-format identities will
    // simply fail at decrypt — better than silently rewriting input.
    input.type = "tel";
    input.placeholder = "+31612345678";
    input.value = attr.v;
    input.addEventListener("input", () => {
      attr.v = input.value;
    });
  } else {
    input.type = "text";
    input.value = attr.v;
    input.addEventListener("input", () => {
      attr.v = input.value;
    });
  }

  inputRow.appendChild(input);

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "pg-policy-attr-delete";
  deleteBtn.setAttribute("aria-label", `Remove ${desc.defaultLabel}`);
  deleteBtn.textContent = "×";
  deleteBtn.addEventListener("click", () => {
    extras.splice(index, 1);
    onDelete();
  });
  inputRow.appendChild(deleteBtn);

  row.appendChild(inputRow);
  return row;
}

function collect(): Policy {
  const result: Policy = {};
  for (const [email, extras] of workingPolicy.entries()) {
    const valid = extras
      .map((a) => ({ t: a.t, v: a.v.trim() }))
      .filter((a) => a.v.length > 0);
    result[email] = [{ t: EMAIL_ATTRIBUTE_TYPE, v: email }, ...valid];
  }
  return result;
}

function ddmmyyyyToHtml(ddmmyyyy: string): string {
  if (!ddmmyyyy) return "";
  const p = ddmmyyyy.split("-");
  return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : "";
}

function htmlToDdmmyyyy(yyyymmdd: string): string {
  if (!yyyymmdd) return "";
  const p = yyyymmdd.split("-");
  return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : "";
}
