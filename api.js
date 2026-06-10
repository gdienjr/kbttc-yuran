// =============================================
// KBTTC API CONFIG
// Tukar API_URL kepada URL Apps Script awak
// =============================================

const API_URL = 'PASTE_YOUR_APPS_SCRIPT_URL_HERE';

async function api(action, params = {}) {
  const body = JSON.stringify({ action, ...params });
  const res = await fetch(API_URL, {
    method: 'POST',
    body,
    headers: { 'Content-Type': 'text/plain' }
  });
  return res.json();
}
