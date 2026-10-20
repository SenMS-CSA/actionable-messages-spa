/** UI wiring for the Actionable Message Sender single-page app. */
(() => {
  const DRAFT_KEY = 'ams.draft';

  const el = {
    accountInfo: document.getElementById('accountInfo'),
    accountEmail: document.getElementById('accountEmail'),
    configBanner: document.getElementById('configBanner'),
    cardJson: document.getElementById('cardJson'),
    cardStatus: document.getElementById('cardStatus'),
    templateVars: document.getElementById('templateVars'),
    sampleSelect: document.getElementById('sampleSelect'),
    fileInput: document.getElementById('fileInput'),
    formatBtn: document.getElementById('formatBtn'),
    from: document.getElementById('fromMailbox'),
    to: document.getElementById('toRecipients'),
    cc: document.getElementById('ccRecipients'),
    subject: document.getElementById('subject'),
    fallbackText: document.getElementById('fallbackText'),
    originator: document.getElementById('originator'),
    saveToSentItems: document.getElementById('saveToSentItems'),
    sendBtn: document.getElementById('sendBtn'),
    draftBtn: document.getElementById('draftBtn'),
    selfTestBtn: document.getElementById('selfTestBtn'),
    previewBtn: document.getElementById('previewBtn'),
    log: document.getElementById('log'),
    previewDialog: document.getElementById('previewDialog'),
    previewContent: document.getElementById('previewContent'),
    copyPreviewBtn: document.getElementById('copyPreviewBtn'),
    closePreviewBtn: document.getElementById('closePreviewBtn')
  };

  let serverStatus = { configured: false };

  function log(message, level, link) {
    const entry = document.createElement('span');
    entry.className = 'log-entry log-' + (level || 'info');
    entry.textContent = '[' + new Date().toLocaleTimeString() + '] ' + message;
    if (link) {
      const anchor = document.createElement('a');
      anchor.href = link.url;
      anchor.target = '_blank';
      anchor.rel = 'noopener noreferrer';
      anchor.textContent = ' ' + link.text;
      entry.appendChild(anchor);
    }
    el.log.appendChild(entry);
    el.log.scrollTop = el.log.scrollHeight;
  }

  function describeError(error) {
    let text = error && error.message ? error.message : String(error);
    if (error && error.requestId) {
      text += ' (Graph request-id: ' + error.requestId + ')';
    }
    return text;
  }

  function portalBladeUrl(blade) {
    return (
      'https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/' +
      (blade || 'Overview') +
      '/appId/' +
      encodeURIComponent(serverStatus.clientId || '')
    );
  }

  const ERROR_HINTS = [
    {
      code: 'invalid_client',
      hint: 'Fix: the client secret is wrong or expired. Create a new secret value (not the secret ID) and update CLIENT_SECRET in .env.',
      portal: true,
      blade: 'Credentials'
    },
    {
      code: 'AADSTS7000215',
      hint: 'Fix: invalid client secret. Use the secret Value, not the secret ID, in CLIENT_SECRET.',
      portal: true,
      blade: 'Credentials'
    },
    {
      code: 'AADSTS7000222',
      hint: 'Fix: the client secret has expired. Create a new one and update CLIENT_SECRET in .env.',
      portal: true,
      blade: 'Credentials'
    },
    {
      code: 'AADSTS700016',
      hint: 'Fix: the client ID was not found in this tenant. Check CLIENT_ID and TENANT_ID in .env.'
    },
    {
      code: 'AADSTS90002',
      hint: 'Fix: the tenant was not found. Check TENANT_ID in .env.'
    },
    {
      code: 'ErrorAccessDenied',
      hint:
        'Fix: the app lacks permission for this mailbox. Add the Microsoft Graph APPLICATION permission ' +
        'Mail.Send (and Mail.ReadWrite for drafts), grant admin consent, and check any Exchange ' +
        'application access policy or RBAC scope. Changes can take a few minutes to apply.',
      portal: true,
      blade: 'CallAnAPI'
    },
    {
      code: 'ErrorInvalidUser',
      hint: 'Fix: the From mailbox does not exist in this tenant.'
    },
    {
      code: 'ResourceNotFound',
      hint: 'Fix: the From mailbox was not found. Use the full user principal name or object ID.'
    },
    {
      code: 'MailboxNotEnabledForRESTAPI',
      hint: 'Fix: the From account has no Exchange Online mailbox, so it cannot send mail.'
    },
    {
      code: 'SenderNotAllowed',
      hint: 'Fix: add the sender to ALLOWED_SENDERS in .env, or clear that setting, then restart the server.'
    }
  ];

  /** Turns a raw token or Graph failure into a one-line remediation plus a portal deep link. */
  function logAuthHint(error) {
    const text = (error && (error.errorMessage || error.message)) || String(error);
    const match = ERROR_HINTS.find((candidate) => text.includes(candidate.code));
    if (!match) {
      return;
    }
    log(
      match.hint,
      'warn',
      match.portal && serverStatus.clientId
        ? {
            text: 'Open the ' + (match.blade || 'Overview') + ' blade →',
            url: portalBladeUrl(match.blade)
          }
        : null
    );
  }

  /* ---------------------------------------------------------------- state */

  function saveDraft() {
    const draft = {
      cardJson: el.cardJson.value,
      templateVars: el.templateVars.value,
      from: el.from.value,
      to: el.to.value,
      cc: el.cc.value,
      subject: el.subject.value,
      fallbackText: el.fallbackText.value,
      originator: el.originator.value,
      saveToSentItems: el.saveToSentItems.checked
    };
    localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
  }

  function restoreDraft() {
    let draft;
    try {
      draft = JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null');
    } catch {
      return;
    }
    if (!draft) {
      return;
    }
    el.cardJson.value = draft.cardJson || '';
    el.templateVars.value = draft.templateVars || '';
    el.from.value = draft.from || '';
    el.to.value = draft.to || '';
    el.cc.value = draft.cc || '';
    el.subject.value = draft.subject || '';
    el.fallbackText.value = draft.fallbackText || '';
    el.originator.value = draft.originator || '';
    el.saveToSentItems.checked = draft.saveToSentItems !== false;
  }

  /* --------------------------------------------------------------- template */

  function currentBuildOptions() {
    return {
      variables: CardBuilder.parseVariables(el.templateVars.value),
      originator: el.originator.value,
      fallbackText: el.fallbackText.value
    };
  }

  function validateTemplate(verbose) {
    const raw = el.cardJson.value.trim();
    if (!raw) {
      setCardStatus('No template loaded', 'pill-muted');
      return null;
    }
    try {
      const result = CardBuilder.build(raw, currentBuildOptions());
      setCardStatus(CardBuilder.describe(result.cardType) + ' · valid', 'pill-ok');
      if (verbose) {
        result.warnings.forEach((warning) => log(warning, 'warn'));
      }
      return result;
    } catch (error) {
      setCardStatus('Invalid: ' + error.message, 'pill-err');
      if (verbose) {
        log(error.message, 'err');
      }
      return null;
    }
  }

  function setCardStatus(text, className) {
    el.cardStatus.textContent = text;
    el.cardStatus.className = 'pill ' + className;
  }

  function formatTemplate() {
    const raw = el.cardJson.value.trim();
    if (!raw) {
      return;
    }
    try {
      el.cardJson.value = JSON.stringify(JSON.parse(raw), null, 2);
      saveDraft();
      validateTemplate(false);
      log('Template formatted.', 'info');
    } catch (error) {
      log('Cannot format: ' + error.message, 'err');
    }
  }

  async function loadSample(path) {
    try {
      const response = await fetch(path, { cache: 'no-cache' });
      if (!response.ok) {
        throw new Error(response.status + ' ' + response.statusText);
      }
      el.cardJson.value = JSON.stringify(await response.json(), null, 2);
      if (!el.subject.value) {
        el.subject.value = 'Action required';
      }
      saveDraft();
      validateTemplate(true);
      log('Loaded sample ' + path, 'ok');
    } catch (error) {
      log('Could not load sample: ' + error.message, 'err');
    }
  }

  /* ---------------------------------------------------------------- server */

  function renderServerState() {
    const ready = serverStatus.configured;
    el.configBanner.classList.toggle('hidden', ready);
    el.accountInfo.classList.toggle('hidden', !ready);
    el.accountEmail.textContent = ready ? 'Client ' + serverStatus.clientId : '';
    [el.sendBtn, el.draftBtn, el.selfTestBtn].forEach((button) => {
      button.disabled = !ready;
    });
  }

  async function initServer() {
    try {
      serverStatus = await Graph.getStatus();
      if (serverStatus.configured && !el.from.value.trim() && serverStatus.defaultSender) {
        el.from.value = serverStatus.defaultSender;
      }
      if (serverStatus.configured && serverStatus.allowedSenders.length) {
        log('Sending is restricted to: ' + serverStatus.allowedSenders.join(', '), 'info');
      }
    } catch (error) {
      serverStatus = { configured: false };
      log('Could not reach the local server: ' + describeError(error), 'err');
    }
    renderServerState();
  }

  /* ------------------------------------------------------------------ send */

  function collectSendOptions(overrideTo) {
    const built = validateTemplate(true);
    if (!built) {
      throw new Error('Fix the card template before sending.');
    }
    const to = overrideTo || el.to.value;
    if (!el.from.value.trim()) {
      throw new Error('A From mailbox is required.');
    }
    if (!Graph.toRecipientList(to).length) {
      throw new Error('At least one recipient is required.');
    }
    if (!el.subject.value.trim()) {
      throw new Error('A subject is required.');
    }
    return {
      from: el.from.value.trim(),
      to,
      cc: overrideTo ? '' : el.cc.value,
      subject: el.subject.value.trim(),
      html: built.html,
      saveToSentItems: el.saveToSentItems.checked
    };
  }

  async function withBusyButtons(action) {
    const buttons = [el.sendBtn, el.draftBtn, el.selfTestBtn];
    buttons.forEach((button) => {
      button.disabled = true;
    });
    try {
      await action();
    } finally {
      renderServerState();
    }
  }

  function handleSend(overrideTo) {
    return withBusyButtons(async () => {
      let options;
      try {
        options = collectSendOptions(overrideTo);
      } catch (error) {
        log(error.message, 'err');
        return;
      }
      try {
        const message = await Graph.sendMail(options);
        const addresses = message.toRecipients.map((r) => r.emailAddress.address).join(', ');
        log('Message sent as ' + options.from + ' to ' + addresses + '.', 'ok');
      } catch (error) {
        log('Send failed: ' + describeError(error), 'err');
        logAuthHint(error);
      }
    });
  }

  function handleDraft() {
    return withBusyButtons(async () => {
      let options;
      try {
        options = collectSendOptions(null);
      } catch (error) {
        log(error.message, 'err');
        return;
      }
      try {
        const draft = await Graph.createDraft(options);
        log('Draft created in ' + options.from + ' (id ' + draft.id.slice(0, 12) + '…).', 'ok');
      } catch (error) {
        log('Draft creation failed: ' + describeError(error), 'err');
        logAuthHint(error);
      }
    });
  }

  function handleSelfTest() {
    const sender = el.from.value.trim();
    if (!sender.includes('@')) {
      log('Enter the From mailbox as an email address to send a test to it.', 'err');
      return;
    }
    log('Sending a test card to ' + sender + '…', 'info');
    return handleSend(sender);
  }

  function handlePreview() {
    const built = validateTemplate(true);
    if (!built) {
      return;
    }
    el.previewContent.textContent = built.html;
    el.previewDialog.showModal();
  }

  /* ------------------------------------------------------------------ init */

  function debounce(fn, delay) {
    let handle;
    return (...args) => {
      clearTimeout(handle);
      handle = setTimeout(() => fn(...args), delay);
    };
  }

  function wireEvents() {
    el.formatBtn.addEventListener('click', formatTemplate);
    el.previewBtn.addEventListener('click', handlePreview);
    el.sendBtn.addEventListener('click', () => handleSend(null));
    el.draftBtn.addEventListener('click', handleDraft);
    el.selfTestBtn.addEventListener('click', handleSelfTest);

    el.closePreviewBtn.addEventListener('click', () => el.previewDialog.close());
    el.copyPreviewBtn.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(el.previewContent.textContent);
        log('HTML body copied to the clipboard.', 'ok');
      } catch (error) {
        log('Copy failed: ' + error.message, 'err');
      }
    });

    el.sampleSelect.addEventListener('change', (event) => {
      if (event.target.value) {
        loadSample(event.target.value);
        event.target.value = '';
      }
    });

    el.fileInput.addEventListener('change', async (event) => {
      const file = event.target.files && event.target.files[0];
      if (!file) {
        return;
      }
      el.cardJson.value = await file.text();
      event.target.value = '';
      saveDraft();
      validateTemplate(true);
      log('Loaded ' + file.name, 'ok');
    });

    const onInput = debounce(() => {
      saveDraft();
      validateTemplate(false);
    }, 300);

    [
      el.cardJson,
      el.templateVars,
      el.from,
      el.to,
      el.cc,
      el.subject,
      el.fallbackText,
      el.originator
    ].forEach((field) => field.addEventListener('input', onInput));
    el.saveToSentItems.addEventListener('change', saveDraft);
  }

  function start() {
    restoreDraft();
    wireEvents();
    validateTemplate(false);
    initServer();
    log('Ready. Enter the mailbox to send as, then send a card.', 'info');
  }

  start();
})();
