/**
 * SVAADH CUSTOM MODALS
 * Replacement for alert() and confirm()
 */

const sAlert = (msg, title = "Svaadh Kitchen", icon = "ℹ️", subtitle = "") => {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 's-modal-overlay';
    overlay.innerHTML = `
      <div class="s-modal">
        <div class="s-modal-icon">${icon}</div>
        <div class="s-modal-title">${title}</div>
        <div class="s-modal-msg">${msg}</div>
        ${subtitle ? `<div style="font-size:0.75rem; color:#888; font-style:italic; margin-top:-15px; margin-bottom:20px;">${subtitle}</div>` : ''}
        <div class="s-modal-footer">
          <button class="s-btn s-btn-primary" id="sModalOk">OK</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    document.getElementById('sModalOk').addEventListener('click', () => {
      document.body.removeChild(overlay);
      resolve(true);
    });
  });
};

const sConfirm = (msg, title = "Confirm Action", icon = "⚠️", subtitle = "") => {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 's-modal-overlay';
    overlay.innerHTML = `
      <div class="s-modal">
        <div class="s-modal-icon">${icon}</div>
        <div class="s-modal-title">${title}</div>
        <div class="s-modal-msg">${msg}</div>
        ${subtitle ? `<div style="font-size:0.75rem; color:#888; font-style:italic; margin-top:-15px; margin-bottom:20px;">${subtitle}</div>` : ''}
        <div class="s-modal-footer">
          <button class="s-btn s-btn-secondary" id="sModalCancel">Cancel</button>
          <button class="s-btn s-btn-primary" id="sModalConfirm">Confirm</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    document.getElementById('sModalCancel').addEventListener('click', () => {
      document.body.removeChild(overlay);
      resolve(false);
    });

    document.getElementById('sModalConfirm').addEventListener('click', () => {
      document.body.removeChild(overlay);
      resolve(true);
    });
  });
};

/**
 * Global Loading Overlay
 * @param {boolean} show - Toggle visibility
 * @param {string} msg - Optional message to show
 */
const sLoading = (show, msg = "Processing...") => {
  const existing = document.getElementById('sLoadingOverlay');
  if (!show) {
    if (existing) document.body.removeChild(existing);
    return;
  }
  if (existing) {
    if (msg) existing.querySelector('.s-loading-msg').textContent = msg;
    return;
  }
  const overlay = document.createElement('div');
  overlay.id = 'sLoadingOverlay';
  overlay.className = 's-loading-overlay';
  overlay.innerHTML = `
    <div class="s-spinner"></div>
    <div class="s-loading-msg">${msg}</div>
  `;
  document.body.appendChild(overlay);
};

/**
 * Button-specific Loading State
 * @param {HTMLElement} btn - The button element
 * @param {boolean} isLoading - Toggle loading state
 * @param {string} loadingText - Text to show while loading
 */
const sBtnLoading = (btn, isLoading, loadingText = "") => {
  if (!btn) return;
  if (isLoading) {
    btn.disabled = true;
    btn.dataset.origHtml = btn.innerHTML;
    const txt = loadingText || btn.innerText;
    btn.innerHTML = `<div class="s-spinner"></div> <span>${txt}...</span>`;
  } else {
    btn.disabled = false;
    if (btn.dataset.origHtml) btn.innerHTML = btn.dataset.origHtml;
  }
};

// Also expose as window globals if needed immediately
window.sAlert = sAlert;
window.sConfirm = sConfirm;
