// Minimal MIME helpers to detect PostGuard ciphertext and pull headers
// out of decrypted plaintext MIME.
//
// We deliberately do NOT try to be a full RFC 5322 parser — only the
// pieces PostGuard needs: thread-related headers and the ASCII-armored
// ciphertext block.

const ARMOR_BEGIN = "-----BEGIN POSTGUARD MESSAGE-----";
const ARMOR_END = "-----END POSTGUARD MESSAGE-----";

export function extractArmoredCiphertext(htmlOrText: string): string | null {
  if (!htmlOrText) return null;
  const begin = htmlOrText.indexOf(ARMOR_BEGIN);
  if (begin < 0) return null;
  const end = htmlOrText.indexOf(ARMOR_END, begin);
  if (end < 0) return null;
  const block = htmlOrText.slice(begin + ARMOR_BEGIN.length, end);
  // Strip whitespace and any HTML tags that may have been wrapped around it.
  return block.replace(/<[^>]+>/g, "").replace(/\s+/g, "");
}

export function looksLikePostGuard(htmlOrText: string): boolean {
  if (!htmlOrText) return false;
  return htmlOrText.indexOf(ARMOR_BEGIN) >= 0;
}

// Pull a single header value (case-insensitive) out of a raw MIME blob.
export function readMimeHeader(rawMime: string, name: string): string | undefined {
  if (!rawMime) return undefined;
  const lcName = name.toLowerCase();
  // Header section ends at the first blank line (CRLF or LF).
  const headerEnd = rawMime.search(/\r?\n\r?\n/);
  const headerSection = headerEnd >= 0 ? rawMime.slice(0, headerEnd) : rawMime;
  // Unfold continuations.
  const unfolded = headerSection.replace(/\r?\n[ \t]+/g, " ");
  for (const line of unfolded.split(/\r?\n/)) {
    const idx = line.indexOf(":");
    if (idx < 0) continue;
    if (line.slice(0, idx).trim().toLowerCase() === lcName) {
      return line.slice(idx + 1).trim();
    }
  }
  return undefined;
}

// Strip MIME headers, return a best-effort body. If multipart we just
// return the original — the SDK normally yields a single text/html part.
export function bodyFromMime(rawMime: string): string {
  const headerEnd = rawMime.search(/\r?\n\r?\n/);
  return headerEnd >= 0 ? rawMime.slice(headerEnd).replace(/^\r?\n\r?\n/, "") : rawMime;
}

// Returns true if the message looks like a multipart/* body.
export function isMultipart(rawMime: string): boolean {
  const ct = readMimeHeader(rawMime, "Content-Type") ?? "";
  return /^multipart\//i.test(ct);
}

export const POSTGUARD_ENCRYPTED_FILENAME = "postguard.encrypted";
export const POSTGUARD_HEADER = "x-postguard";
export const POSTGUARD_HEADER_VALUE = "0.1.0";
