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

  // ── Constants ─────────────────────────────────────────────────
  const ALLOWED_CHARS         = /^[0-9*#+]$/;
  const ALLOWED_CHARS_GLOBAL  = /[^0-9*#+]/g;  // Used to strip invalid chars from pasted input
  const MAX_LENGTH            = 30;

  // ── Utility: Set status message ───────────────────────────────
  function setStatus(message, type = '') {
    statusMsg.textContent = message;
    statusMsg.className   = `status-msg ${type}`;
    if (message) {
      setTimeout(() => {
        statusMsg.textContent = '';
        statusMsg.className   = 'status-msg';
      }, 4000);
    }
  }

  // ── Utility: Sanitize a raw string to only allowed characters ─
  function sanitizeInput(raw) {
    return raw.replace(ALLOWED_CHARS_GLOBAL, '');
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

    const encodedNumber = encodeURIComponent(sanitized);
    const ciscotelUrl   = `CISCOTEL:${encodedNumber}`;

    try {
      const anchor         = document.createElement('a');
      anchor.href          = ciscotelUrl;
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

  // ── Event: Paste handler ──────────────────────────────────────
  // Intercepts paste, strips all non-dialable characters, and
  // inserts only the sanitized result — respecting MAX_LENGTH.
  dialDisplay.addEventListener('paste', (event) => {
    event.preventDefault();   // Block the default paste behaviour

    const raw       = (event.clipboardData || window.clipboardData).getData('text');
    const cleaned   = sanitizeInput(raw);

    if (!cleaned) {
      setStatus('Clipboard content contains no dialable characters.', 'error');
      return;
    }

    // Respect max length: only take as many chars as still fit
    const remaining     = MAX_LENGTH - dialDisplay.value.length;
    const toInsert      = cleaned.slice(0, remaining);
    const wasClipped    = cleaned.length > remaining;

    // Insert at the current cursor position rather than always appending
    const start = dialDisplay.selectionStart;
    const end   = dialDisplay.selectionEnd;
    const current = dialDisplay.value;

    dialDisplay.value = current.slice(0, start) + toInsert + current.slice(end);

    // Restore cursor position after the inserted text
    const newCursorPos = start + toInsert.length;
    dialDisplay.setSelectionRange(newCursorPos, newCursorPos);

    if (wasClipped) {
      setStatus(`Number trimmed to ${MAX_LENGTH} digits maximum.`, 'error');
    } else {
      setStatus(`Pasted: ${toInsert}`, 'success');
    }

    console.log(`[DialPad] Pasted raw: "${raw}" → sanitized: "${toInsert}"`);
  });

  // ── Event: Direct typing / input sanitization ─────────────────
  // Since readonly is removed, users can also type directly into
  // the field. This ensures only allowed characters are accepted.
  dialDisplay.addEventListener('keydown', (event) => {
    const key = event.key;

    // Allow control keys to function normally
    const controlKeys = [
      'Backspace', 'Delete', 'ArrowLeft', 'ArrowRight',
      'ArrowUp', 'ArrowDown', 'Tab', 'Enter', 'Home', 'End'
    ];

    if (controlKeys.includes(key)) {
      // Handle Enter to dial
      if (key === 'Enter') {
        event.preventDefault();
        invokeCiscoTel(dialDisplay.value);
      }
      // Allow all other control keys (Backspace, arrows, etc.) to work natively
      return;
    }

    // Allow Ctrl+A, Ctrl+C, Ctrl+V, Ctrl+X for clipboard operations
    if (event.ctrlKey || event.metaKey) return;

    // Block any character that is not in the allowed set
    if (!ALLOWED_CHARS.test(key)) {
      event.preventDefault();
      setStatus(`"${key}" is not a valid dial character.`, 'error');
      return;
    }

    // Enforce max length for direct typing
    if (dialDisplay.value.length >= MAX_LENGTH) {
      event.preventDefault();
      setStatus('Maximum number length reached.', 'error');
    }
  });

})();
