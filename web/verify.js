/**
 * Certificate Verification System - Frontend JavaScript
 * 
 * This script handles the verification of certificates by:
 * 1. Loading certificate data and hash index from JSON files
 * 2. Verifying certificate hashes against the blockchain-like chain
 * 3. Displaying verification results with certificate details
 * 4. Providing download and sharing functionality
 */

/**
 * Loads certificate data and hash index from JSON files
 * @returns {Promise<Object>} Object containing certificates array and hash lookup index
 */
async function loadData() {
  const [certsRes, indexRes] = await Promise.all([
    fetch('data/certs.json'),
    fetch('data/hash_index.json'),
  ]);
  if (!certsRes.ok || !indexRes.ok) {
    throw new Error('Failed to load verification data');
  }
  const certs = await certsRes.json();
  const index = await indexRes.json();
  return { certs, index };
}

/**
 * Computes SHA-256 hash of a string and returns it as hexadecimal
 * @param {string} str - Input string to hash
 * @returns {Promise<string>} Hexadecimal SHA-256 hash
 */
async function sha256Hex(str) {
  const enc = new TextEncoder();
  const data = enc.encode(str);
  const digest = await crypto.subtle.digest('SHA-256', data);
  const bytes = new Uint8Array(digest);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    hex += bytes[i].toString(16).padStart(2, '0');
  }
  return hex;
}

/**
 * Combines certificate fields into a single string for hashing
 * @param {Object} rec - Certificate record
 * @returns {string} Concatenated fields string
 */
function concatFields(rec) {
  return `${rec.CertID}|${rec.RecipientName}|${rec.CourseTitle}|${rec.DateIssued}|${rec.PreviousHash}`;
}

/**
 * Validates that a string is a 64-character hexadecimal value (SHA-256 hash)
 * @param {string} s - String to validate
 * @returns {boolean} True if string is a valid 64-character hex value
 */
function isHex64(s) {
  const clean = (s || '').trim();
  return /^[0-9a-f]{64}$/.test(clean);
}

/**
 * Updates the result container with HTML content
 * @param {HTMLElement} el - Element to update
 * @param {string} content - HTML content to insert
 */
function renderResult(el, content) {
  el.innerHTML = content;
}

/**
 * Formats a long hash for display by truncating the middle section
 * @param {string} hash - Hash to format
 * @param {number} visibleStart - Number of characters to show at start
 * @param {number} visibleEnd - Number of characters to show at end
 * @returns {string} Formatted hash with ellipsis in the middle
 */
function formatHash(hash, visibleStart = 10, visibleEnd = 6) {
  const clean = (hash || '').trim();
  if (!clean) return '';
  if (clean.length <= visibleStart + visibleEnd + 3) return clean;
  return `${clean.slice(0, visibleStart)}...${clean.slice(-visibleEnd)}`;
}

/**
 * Downloads a certificate PDF by CertID or falls back to printing the current page
 * @param {string} certId - Certificate ID to download
 * @returns {Promise<void>}
 */
async function downloadCertificateById(certId) {
  const safeId = encodeURIComponent(certId || '');
  const apiUrl = `/api/download-certificate/${safeId}`;
  try {
    // Create and trigger a download link to the API endpoint
    const a = document.createElement('a');
    a.href = apiUrl;
    a.download = '';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (e) {
    // Fallback: print current result card as PDF if download fails
    window.print();
  }
}

function renderSuccess(el, rec, extraBlock = '') {
  const prevLink = rec.PreviousHash ? `?hash=${rec.PreviousHash}` : '';
  const qrHtml = rec.QRCodeURL || rec.QRCodePath ? `
    <div class="qr">
      ${rec.QRCodeURL ? `<img src="${rec.QRCodeURL}" alt="QR code to verify certificate" />` : `<img src="${rec.QRCodePath}" alt="QR code to verify certificate" />`}
      <div class="qr-meta">
        <div><strong>Scan QR</strong> to open verification page.</div>
        ${rec.VerificationURL ? `<div><a href="${rec.VerificationURL}" target="_blank" rel="noopener">Open Verification Page</a></div>` : ''}
      </div>
    </div>
  ` : '';
  const actionsHtml = `
    <div class="actions">
      ${rec.VerificationURL ? `<button class="btn" onclick="navigator.clipboard && navigator.clipboard.writeText('${rec.VerificationURL}').then(()=>{const m=document.getElementById('copyMsg'); if(m){m.textContent='Copied!'; setTimeout(()=>m.textContent='', 1500);} })">Copy verification link</button>` : ''}
      <button class="btn" onclick="downloadCertificateById('${rec.CertID || ''}')">Download certificate</button>
      <span id="copyMsg" class="copy-msg" aria-live="polite"></span>
    </div>
  `;
  el.innerHTML = `
    <div class="card">
      <div class="status ok"><strong>Verified:</strong> Certificate is valid.</div>
      <div class="grid">
        <div class="details">
          <dl class="dl">
            <div><dt>CertID</dt><dd>${rec.CertID}</dd></div>
            <div><dt>Recipient</dt><dd>${rec.RecipientName}</dd></div>
            <div><dt>Course</dt><dd>${rec.CourseTitle}</dd></div>
            <div><dt>Date Issued</dt><dd>${rec.DateIssued}</dd></div>
            <div><dt>PreviousHash</dt><dd><div class="hash-row"><span class="hash">${formatHash(rec.PreviousHash)}</span><button class="button secondary" onclick="navigator.clipboard && navigator.clipboard.writeText('${rec.PreviousHash}').then(()=>{this.textContent='Copied!'; setTimeout(()=>this.textContent='Copy', 1500);})">Copy</button></div></dd></div>
            <div><dt>CurrentHash</dt><dd><div class="hash-row"><span class="hash">${formatHash(rec.CurrentHash)}</span><button class="button secondary" onclick="navigator.clipboard && navigator.clipboard.writeText('${rec.CurrentHash}').then(()=>{this.textContent='Copied!'; setTimeout(()=>this.textContent='Copy', 1500);})">Copy</button></div></dd></div>
          </dl>
          ${prevLink ? `<a class="prev-link" href="${prevLink}">Verify previous certificate</a>` : ''}
        </div>
        ${qrHtml}
      </div>
      ${actionsHtml}
      ${extraBlock}
    </div>
  `;
}

function renderFailure(el, message, rec) {
  el.innerHTML = `
    <div class="card">
      <div class="status fail"><strong>Verification failed:</strong> ${message}</div>
      ${rec ? `
        <dl class="dl">
          <div><dt>CertID</dt><dd>${rec.CertID}</dd></div>
          <div><dt>Recipient</dt><dd>${rec.RecipientName}</dd></div>
          <div><dt>Course</dt><dd>${rec.CourseTitle}</dd></div>
          <div><dt>Date Issued</dt><dd>${rec.DateIssued}</dd></div>
          <div><dt>PreviousHash</dt><dd><div class="hash-row"><span class="hash">${formatHash(rec.PreviousHash)}</span>${rec.PreviousHash ? `<button class="button secondary" onclick="navigator.clipboard && navigator.clipboard.writeText('${rec.PreviousHash}').then(()=>{this.textContent='Copied!'; setTimeout(()=>this.textContent='Copy', 1300);})">Copy</button>` : ''}</div></dd></div>
          <div><dt>Stored CurrentHash</dt><dd><div class="hash-row"><span class="hash">${formatHash(rec.CurrentHash)}</span><button class="button secondary" onclick="navigator.clipboard && navigator.clipboard.writeText('${rec.CurrentHash}').then(()=>{this.textContent='Copied!'; setTimeout(()=>this.textContent='Copy', 1300);})">Copy</button></div></dd></div>
        </dl>
      ` : ''}
    </div>
  `;
}

/**
 * Verifies a certificate by its hash against the certificate chain
 * @param {string} hash - SHA-256 hash of the certificate to verify
 * @param {Object} data - Object containing certificates and hash index
 * @param {HTMLElement} el - Element to render results into
 * @returns {Promise<void>}
 */
/**
 * Verifies a certificate by its hash against the certificate chain
 * @param {string} hash - SHA-256 hash of the certificate to verify
 * @param {Object} data - Object containing certificates and hash index
 * @param {HTMLElement} el - Element to render results into
 * @returns {Promise<void>}
 */
async function verifyHash(hash, data, el) {
  const clean = (hash || '').trim().toLowerCase();
  if (!clean) {
    renderFailure(el, 'Please paste a hash to verify.');
    return;
  }
  if (!isHex64(clean)) {
    renderFailure(el, 'Hash must be 64 hex characters (0-9, a-f).');
    return;
  }
  const idx = data.index[clean];
  if (idx === undefined) {
    renderFailure(el, 'No certificate found with the provided hash.');
    return;
  }
  const rec = data.certs[idx];
  const concat = concatFields(rec);
  const recomputed = await sha256Hex(concat);
  if (recomputed === rec.CurrentHash) {
    let extra = '';
    const genesis = '0'.repeat(64);
    if (rec.PreviousHash && rec.PreviousHash !== genesis) {
      const prevIdx = data.index[rec.PreviousHash];
      if (prevIdx === undefined) {
        extra = '<div class="warn">Warning: PreviousHash not found in dataset (chain may be incomplete).</div>';
      } else {
        const prevRec = data.certs[prevIdx];
        if (prevRec.CurrentHash !== rec.PreviousHash) {
          extra = '<div class="fail">Chain link mismatch: previous record hash differs.</div>';
        }
      }
    }
    renderSuccess(el, rec, extra);
  } else {
    renderFailure(el, 'Recomputed hash does not match stored CurrentHash.', rec);
  }
}

(async function init() {
  const hashInput = document.getElementById('hashInput');
  const verifyBtn = document.getElementById('verifyBtn');
  const resultEl = document.getElementById('result');

  let data;
  try {
    data = await loadData();
  } catch (e) {
    renderFailure(resultEl, 'Could not load verification data.');
    console.error(e);
    return;
  }

  verifyBtn.addEventListener('click', () => {
    verifyHash(hashInput.value, data, resultEl);
  });

  const urlParams = new URLSearchParams(window.location.search);
  const prefill = urlParams.get('hash');
  if (prefill) {
    hashInput.value = prefill;
    verifyHash(prefill, data, resultEl);
  }
  // Enhance interactions: support Enter key to verify
  // (placed after verifyBtn click binding in init)
})();