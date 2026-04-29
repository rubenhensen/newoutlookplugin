// Inline policy editor used both for "Manage Access" (recipient policies)
// and "Sign" (sender attribute selection). Renders into #view-policy-editor.

import { EMAIL_ATTRIBUTE_TYPE, SUPPORTED_ATTRIBUTES } from "../lib/attributes";
import { Policy, AttributeRequest } from "../lib/types";
import { t } from "../lib/i18n";
import { showView } from "./taskpane";

interface PolicyEditorOptions {
  initialPolicy: Policy;
  sign: boolean;
  onSave: (next: Policy) => void;
  onCancel: () => void;
}

export function openPolicyEditor(opts: PolicyEditorOptions): void {
  const titleEl = document.getElementById("pg-policy-title")!;
  const recipientsEl = document.getElementById("pg-policy-recipients")!;
  const saveBtn = document.getElementById("pg-policy-save") as HTMLButtonElement;
  const cancelBtn = document.getElementById("pg-policy-cancel") as HTMLButtonElement;

  titleEl.textContent = opts.sign ? t("policyEditorTitleSign") : t("policyEditorTitle");
  saveBtn.textContent = t("policyEditorSave");
  cancelBtn.textContent = t("policyEditorCancel");

  recipientsEl.innerHTML = "";

  for (const [email, attrs] of Object.entries(opts.initialPolicy)) {
    recipientsEl.appendChild(renderRecipient(email, attrs));
  }

  // Replace listeners by cloning the button (cheap way to drop prior bindings).
  const saveClone = saveBtn.cloneNode(true) as HTMLButtonElement;
  saveBtn.replaceWith(saveClone);
  saveClone.addEventListener("click", () => {
    const next = collect(recipientsEl);
    opts.onSave(next);
  });

  const cancelClone = cancelBtn.cloneNode(true) as HTMLButtonElement;
  cancelBtn.replaceWith(cancelClone);
  cancelClone.addEventListener("click", () => opts.onCancel());

  showView("policy_editor");
}

function renderRecipient(email: string, attrs: AttributeRequest[]): HTMLElement {
  const section = document.createElement("div");
  section.className = "pg-policy-recipient";
  section.dataset.email = email;

  const label = document.createElement("div");
  label.className = "pg-policy-recipient-email";
  label.textContent = email;
  section.appendChild(label);

  for (const desc of SUPPORTED_ATTRIBUTES) {
    const isLocked = desc.type === EMAIL_ATTRIBUTE_TYPE;
    const existing = attrs.find((a) => a.t === desc.type);
    const isChecked = isLocked || !!existing;

    const row = document.createElement("div");
    row.className = "pg-policy-attr" + (isLocked ? " locked" : "");

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = isChecked;
    checkbox.dataset.attrType = desc.type;
    checkbox.disabled = isLocked;

    const labelEl = document.createElement("label");
    labelEl.textContent = t(desc.type, desc.defaultLabel);

    const value = document.createElement("input");
    value.type = "text";
    value.placeholder = t(desc.type, desc.defaultLabel);
    value.value = isLocked ? email : existing?.v ?? "";
    value.dataset.attrType = desc.type;
    value.readOnly = isLocked;

    row.appendChild(checkbox);
    row.appendChild(labelEl);
    row.appendChild(value);
    section.appendChild(row);
  }

  return section;
}

function collect(container: HTMLElement): Policy {
  const result: Policy = {};
  const sections = container.querySelectorAll<HTMLElement>(".pg-policy-recipient");
  for (const section of Array.from(sections)) {
    const email = section.dataset.email!;
    const attrs: AttributeRequest[] = [];
    const checkboxes = section.querySelectorAll<HTMLInputElement>(
      'input[type="checkbox"]:checked'
    );
    for (const cb of Array.from(checkboxes)) {
      const type = cb.dataset.attrType!;
      const valueInput = section.querySelector<HTMLInputElement>(
        `input[type="text"][data-attr-type="${type}"]`
      );
      const v = (valueInput?.value ?? "").trim();
      if (!v && type !== EMAIL_ATTRIBUTE_TYPE) continue;
      attrs.push({ t: type, v });
    }
    result[email] = attrs;
  }
  return result;
}
