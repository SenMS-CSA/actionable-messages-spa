/** Client for this app's server-side mail API, which calls Microsoft Graph with application credentials. */
const Graph = (() => {
  async function request(method, path, body) {
    const response = await fetch(path, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body)
    });

    if (!response.ok) {
      throw await toApiError(response);
    }
    return response.json();
  }

  async function toApiError(response) {
    let message = response.status + ' ' + response.statusText;
    let requestId;
    try {
      const payload = await response.json();
      if (payload && payload.error) {
        message = payload.error.code + ': ' + payload.error.message;
        requestId = payload.requestId;
      }
    } catch {
      /* non-JSON error body */
    }
    const error = new Error(message);
    error.status = response.status;
    error.requestId = requestId;
    return error;
  }

  function getStatus() {
    return request('GET', '/api/status');
  }

  function toRecipientList(value) {
    return String(value || '')
      .split(/[;,]/)
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));
  }

  function buildMessage(options) {
    const message = {
      subject: options.subject,
      body: { contentType: 'HTML', content: options.html },
      toRecipients: toRecipientList(options.to)
    };
    const cc = toRecipientList(options.cc);
    if (cc.length) {
      message.ccRecipients = cc;
    }
    return message;
  }

  async function sendMail(options) {
    const message = buildMessage(options);
    if (!message.toRecipients.length) {
      throw new Error('At least one recipient is required.');
    }
    await request('POST', '/api/send', {
      from: options.from,
      message,
      saveToSentItems: options.saveToSentItems !== false
    });
    return message;
  }

  function createDraft(options) {
    return request('POST', '/api/draft', { from: options.from, message: buildMessage(options) });
  }

  return { getStatus, sendMail, createDraft, toRecipientList };
})();
