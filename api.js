// =============================================
// KBTTC API CONFIG
// Tukar API_URL kepada URL Apps Script awak
// =============================================

const API_URL = 'https://script.google.com/macros/s/AKfycbyNWNLOkpBT8LOWQRCF4ZbdfavxEhLlkbeRZqSQ16tAGsn7p5M3l2pULG0arUUAlpi9/exec';

async function api(action, params = {}) {
  const body = JSON.stringify({ action, ...params });
  const res = await fetch(API_URL, {
    method: 'POST',
    body,
    headers: { 'Content-Type': 'text/plain' }
  });
  return res.json();
}
