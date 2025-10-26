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

function concatFields(rec) {
  return `${rec.CertID}|${rec.RecipientName}|${rec.CourseTitle}|${rec.DateIssued}|${rec.PreviousHash}`;
}

function isHex64(s) {
  const clean = (s || '').trim();
  return /^[0-9a-f]{64}$/.test(clean);
}

function renderResult(el, content) {
  el.innerHTML = content;
}

function formatHash(hash, visibleStart = 10, visibleEnd = 6) {
  const clean = (hash || '').trim();
  if (!clean) return '';
  if (clean.length <= visibleStart + visibleEnd + 3) return clean;
  return `${clean.slice(0, visibleStart)}...${clean.slice(-visibleEnd)}`;
}

// Download certificate helper: tries direct file, falls back to print
async function downloadCertificateById(certId) {
  const fileName = certId ? `${certId}.pdf` : 'certificate.pdf';
  const candidates = [
    `data/certificates/${fileName}`
  ];
  for (const url of candidates) {
    try {
      const res = await fetch(url, { method: 'HEAD' });
      if (res.ok) {
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        return;
      }
    } catch (e) {
      // ignore and continue to fallback
    }
  }
  // Fallback: print current result card as PDF
  window.print();
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