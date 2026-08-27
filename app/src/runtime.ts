/** Runtime flags from /api/health — default web-safe for production. */
let webMode = true;

export function setWebMode(on: boolean) {
  webMode = on;
}

export function isWebMode() {
  return webMode;
}
