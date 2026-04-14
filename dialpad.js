(() => {
  'use strict';

  // ── Teams SDK Initialization ─────────────────────────────────
  microsoftTeams.app.initialize().then(() => {
    console.log('[DialPad] Microsoft Teams SDK initialized.');
  }).catch((err) => {
    console.warn('[DialPad] Teams SDK init warning (may run outside Teams):', err);
  });

  // ── DOM References ────────────────────────────────────────────
  const dialDisplay  = document.getElementById('dialDisplay');
  const dialBtn      = document.getElementById('dialBtn');
  const clearBtn     = document.getElementById('clearBtn');
  const dialpadGrid  = document.getElementById('dialpadGrid');
  const statusMsg    = document.getElementById('statusMsg');

  // ── Allowed input characters ──────────────────────────────────
  const ALLOWED_CHARS = /^[0-9*#+]$/;
  const MAX_LENGTH    = 30;

  // ── Utility: Set status message ───────────────────────────────
  function setStatus(message, type = '') {
    statusMsg.textContent  = message;
    statusMsg.className    = `status-msg ${type}`;
    if (message) {
      setTimeout(() => {
        statusMsg.textContent = '';
        statusMsg.className   = 'status-msg';
      }, 4000);
    }
  }

  // ── Append digit to display ───────────────────────────────────
  function appendDigit(digit) {
    if (dialDisplay.value.length >= MAX_LENGTH) {
      setStatus('Maximum number length reached.', 'error');
      return;
    }
    dialDisplay.value += digit;
    setStatus('');
  }

  // ── Delete last character ─────────────────────────────────────
  function deleteLastChar() {
    dialDisplay.value = dialDisplay.value.slice(0, -1);
    setStatus('');
  }

  // ── Invoke CISCOTEL Protocol Handler ─────────────────────────
  function invokeCiscoTel(number) {
    const sanitized = number.trim();

    if (!sanitized) {
      setStatus('Please enter a number to dial.', 'error');
      return;
    }

    // Encode the number safely for use in a URL
    const encodedNumber = encodeURIComponent(sanitized);
    const ciscotelUrl   = `CISCOTEL:${encodedNumber}`;

    try {
      // Use an invisible anchor to trigger the protocol handler
      const anchor    = document.createElement('a');
      anchor.href     = ciscotelUrl;
      anchor.style.display = 'none';
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);

      setStatus(`Calling ${sanitized}…`, 'success');
      console.log(`[DialPad] Invoked: ${ciscotelUrl}`);
    } catch (err) {
      setStatus('Failed to invoke CISCOTEL handler. Is Cisco software installed?', 'error');
      console.error('[DialPad] CISCOTEL invocation error:', err);
    }
  }

  // ── Event: Dial pad button clicks ─────────────────────────────
  dialpadGrid.addEventListener('click', (event) => {
    const btn = event.target.closest('.pad-btn');
    if (!btn) return;
    const value = btn.getAttribute('data-value');
    if (value) appendDigit(value);
  });

  // ── Event: Clear / Backspace button ───────────────────────────
  clearBtn.addEventListener('click', deleteLastChar);

  // ── Event: Dial button ────────────────────────────────────────
  dialBtn.addEventListener('click', () => {
    invokeCiscoTel(dialDisplay.value);
  });

  // ── Event: Physical keyboard input ───────────────────────────
  document.addEventListener('keydown', (event) => {
    const key = event.key;

    if (ALLOWED_CHARS.test(key)) {
      appendDigit(key);
    } else if (key === 'Backspace') {
      deleteLastChar();
    } else if (key === 'Enter') {
      invokeCiscoTel(dialDisplay.value);
    }
  });

})();
