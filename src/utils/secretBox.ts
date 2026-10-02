// Client-side protection for API keys: they are encrypted with AES-GCM using a key derived
// from a passphrase (PBKDF2, 210k iterations). Keys are only readable after the user unlocks
// the session, so a script or extension reading local storage finds only ciphertext.

const ITERATIONS = 210000;

export interface EncryptedBox {
  v: 1;
  /** Random PBKDF2 salt, base64. */
  salt: string;
  /** Random AES-GCM IV, base64. */
  iv: string;
  /** Ciphertext (a JSON map), base64. */
  data: string;
}

/** A byte buffer backed by a plain ArrayBuffer (satisfies BufferSource with DOM types). */
const bytes = (input: Uint8Array): ArrayBuffer => {
  const out = new ArrayBuffer(input.byteLength);
  new Uint8Array(out).set(input);
  return out;
};

const toBase64 = (input: Uint8Array) => btoa(String.fromCharCode(...input));
const fromBase64 = (text: string) => Uint8Array.from(atob(text), ch => ch.charCodeAt(0));

const deriveKey = async (passphrase: string, salt: Uint8Array): Promise<CryptoKey> => {
  const base = await crypto.subtle.importKey('raw', bytes(new TextEncoder().encode(passphrase)), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: bytes(salt), iterations: ITERATIONS, hash: 'SHA-256' },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
};

export const encryptJson = async (value: unknown, passphrase: string): Promise<EncryptedBox> => {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(passphrase, salt);
  const cipher = await crypto.subtle.encrypt({ name: 'AES-GCM', iv: bytes(iv) }, key, bytes(new TextEncoder().encode(JSON.stringify(value))));
  return { v: 1, salt: toBase64(salt), iv: toBase64(iv), data: toBase64(new Uint8Array(cipher)) };
};

export const decryptJson = async <T>(box: EncryptedBox, passphrase: string): Promise<T> => {
  const key = await deriveKey(passphrase, fromBase64(box.salt));
  try {
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: bytes(fromBase64(box.iv)) }, key, bytes(fromBase64(box.data)));
    return JSON.parse(new TextDecoder().decode(plain)) as T;
  } catch {
    throw new Error('Wrong passphrase.');
  }
};
