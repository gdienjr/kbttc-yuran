// =============================================
// KBTTC API CONFIG
// Tukar API_URL kepada URL Apps Script awak
// =============================================

const API_URL = 'https://script.google.com/macros/s/AKfycbwDObNLwvz-1XHLt-Dx6-7Ar5bDc0OY4TE-eVigv27CAUgpk7E019H2kv0gUWdKjMez/exec';

async function api(action, params = {}) {
  const body = JSON.stringify({ action, ...params });
  const res = await fetch(API_URL, {
    method: 'POST',
    body,
    headers: { 'Content-Type': 'text/plain' }
  });
  return res.json();
}
