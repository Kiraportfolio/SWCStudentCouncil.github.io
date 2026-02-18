// ==========================================
// API Client — Shared across all pages
// Communicates with Google Apps Script REST API
// ==========================================

// ===== CONFIGURATION =====
// Set this to your deployed GAS Web App URL
// After deploying GAS, replace this value with the actual URL
const API_BASE = (() => {
  // Check if there's a custom API URL set via <meta> tag or global variable
  if (typeof window !== 'undefined' && window.GAS_API_URL) {
    return window.GAS_API_URL;
  }
  // Check for meta tag: <meta name="gas-api-url" content="https://...">
  const meta = document.querySelector('meta[name="gas-api-url"]');
  if (meta && meta.content && meta.content.indexOf('YOUR_DEPLOYMENT_ID') === -1) {
    return meta.content;
  }
  // Default placeholder — must be replaced before deploying
  return 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
})();

// ===== API FUNCTIONS =====

/**
 * Make a GET request to the GAS API
 * @param {string} action - The API action name (e.g., 'getVoteData')
 * @param {Object} params - Optional query parameters
 * @returns {Promise<Object>} - The JSON response
 */
async function apiGet(action, params = {}) {
  const url = new URL(API_BASE);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([k, v]) => {
    if (v !== undefined && v !== null) url.searchParams.set(k, String(v));
  });

  try {
    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });
    if (!res.ok) throw new Error('API error: ' + res.status);
    return await res.json();
  } catch (err) {
    console.error('[apiGet] ' + action + ' failed:', err);
    throw err;
  }
}

/**
 * Make a POST request to the GAS API
 * @param {string} action - The API action name (e.g., 'submitVote')
 * @param {Object} body - The request body data
 * @returns {Promise<Object>} - The JSON response
 */
async function apiPost(action, body = {}) {
  try {
    const res = await fetch(API_BASE, {
      method: 'POST',
      redirect: 'follow',
      // Use text/plain to avoid CORS preflight (GAS doesn't handle OPTIONS)
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, ...body })
    });
    if (!res.ok) throw new Error('API error: ' + res.status);
    return await res.json();
  } catch (err) {
    console.error('[apiPost] ' + action + ' failed:', err);
    throw err;
  }
}

// ===== NAVIGATION =====

/**
 * Navigate to a page within the app
 * Works for both GitHub Pages and local file serving
 */
function navigateTo(page) {
  if (page === 'index') {
    // Try going to root, fallback to index.html
    const loc = window.location;
    if (loc.protocol === 'file:') {
      window.location.href = 'index.html';
    } else {
      // GitHub Pages — go to root or the base path
      const basePath = window.APP_BASE_PATH || '/';
      window.location.href = basePath;
    }
  } else {
    window.location.href = page + '.html';
  }
}

// ===== UTILITY =====

/**
 * Check if the API is properly configured
 */
function isApiConfigured() {
  return API_BASE && API_BASE.indexOf('YOUR_DEPLOYMENT_ID') === -1;
}

/**
 * Format a number with commas
 */
function formatNumber(n) {
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
